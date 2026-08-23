using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Office.Automation.Com;
using Office.Automation.OpenXml;
using Office.Automation.Runtime;

namespace Office.Automation.Host;

/// <summary>
/// office-rpc/1 dispatcher: handshake, ping and office.command.execute routing
/// (proposal §12). Owns the per-sidecar STA queue and the COM backend.
///
/// Routing (proposal §6 backends):
///   - deck.compile                      → Open XML worker (no Office needed)
///   - document.inspect                  → COM when attached, Open XML (pptx) otherwise
///   - batch.convert / batch.replace_text → COM (per-app: this sidecar handles
///     only its own app's files — §8.2 one process per application)
///   - slide.render                      → COM (PowerPoint only)
///
/// Progressive discovery (§10.2): the handshake capability manifest lists
/// desktop_com capabilities only when the COM backend attached successfully.
/// </summary>
public sealed class CommandRouter : IDisposable
{
    private static readonly TimeSpan AttachBudget = TimeSpan.FromSeconds(60);
    private static readonly CapabilityCatalog Catalog = CapabilityCatalog.Current;

    private readonly string _app;
    private readonly string _workspaceRoot;
    private readonly StaDispatcher _sta = new();
    private readonly OfficeComBackend? _com;
    private readonly TemplateRegistry _templates;
    private readonly object _attachLock = new();
    private readonly HostNotificationQueue _notifications;
    private readonly InMemoryJobTracker _jobs;
    private bool _comAttached;
    private bool _comAttachAttempted;
    private bool _applicationStartedEventPublished;

    public CommandRouter(string app)
        : this(
            app,
            enableDesktopCom: true,
            workspaceRoot: Directory.GetCurrentDirectory(),
            includeDefaultTemplateDirectories: true)
    {
    }

    internal CommandRouter(
        string app,
        bool enableDesktopCom,
        string? workspaceRoot = null,
        IEnumerable<string>? templateDirectories = null,
        bool includeDefaultTemplateDirectories = false)
    {
        _app = app;
        _workspaceRoot = WorkspaceGuard.CanonicalizeRoot(
            workspaceRoot ?? Directory.GetCurrentDirectory());
        _templates = new TemplateRegistry(
            templateDirectories,
            includeDefaultTemplateDirectories);
        _notifications = new HostNotificationQueue(app, HostId);
        _jobs = new InMemoryJobTracker(_notifications);
        _com = enableDesktopCom && ComBackendFactory.IsSupported(app)
            ? ComBackendFactory.Create(app, _sta)
            : null;
    }

    public string App => _app;

    public bool ComAttached => _comAttached;

    internal IReadOnlyList<string> DrainNotifications() => _notifications.Drain();

    /// <summary>Set by office.host.shutdown: the pipe loop exits after the current connection.</summary>
    public bool ShutdownRequested { get; private set; }

    /// <summary>Attaches the §27 criterion-10 audit trail to a command result.</summary>
    private object AddAudit(object result, CommandPolicy policy, long elapsedMs)
    {
        var node = JsonSerializer.SerializeToNode(result)!.AsObject();
        OfficeSecurityPosture? posture = _comAttached ? _com?.SecurityPosture : null;
        var audit = new JsonObject
        {
            ["policy"] = policy.EffectivePolicy.DeepClone(),
            ["confirmation"] = policy.ConfirmationForAudit?.DeepClone(),
            ["security"] = new JsonObject
            {
                ["automation_security"] = new JsonObject
                {
                    ["applicable"] = posture is not null,
                    ["observed"] = posture?.AutomationSecurity,
                    ["expected"] = posture is null ? null : 3,
                    ["enforced"] = posture?.AutomationSecurity == 3,
                },
                ["display_alerts_disabled"] = posture?.DisplayAlertsDisabled,
                ["execute_mso"] = "deny_all",
                ["external_links_auto_update_disabled"] =
                    posture?.ExternalLinksAutoUpdateDisabled,
                ["macros"] = posture is null ? "not_applicable" : "deny_observed",
                ["workspace_only"] = true,
            },
            ["backend"] = node.TryGetPropertyValue("backend", out var backendNode)
                ? backendNode?.DeepClone()
                : null,
            ["host_version"] = HostBuildInfo.Version,
            ["application"] = _comAttached
                ? JsonSerializer.SerializeToNode(_com!.GetApplicationInfo())
                : null,
            ["duration_ms"] = elapsedMs,
        };
        node["audit"] = audit;
        return node;
    }

    /// <summary>Processes one JSON-RPC request line and returns the response line.</summary>
    public string Dispatch(string requestJson)
    {
        JsonElement id = JsonSerializer.SerializeToElement<object?>(null);
        string method;
        CommandPolicy? commandPolicy = null;
        RpcMethodDefinition? rpcMethod = null;
        try
        {
            using var request = JsonDocument.Parse(requestJson);
            var root = request.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                throw new OfficeArgumentException("JSON-RPC request must be an object");
            }
            id = root.TryGetProperty("id", out var idElement) ? idElement.Clone() : id;
            method = root.TryGetProperty("method", out var methodElement)
                && methodElement.ValueKind == JsonValueKind.String
                ? methodElement.GetString()!
                : throw new OfficeArgumentException("JSON-RPC method must be a string");
            var parameters = root.TryGetProperty("params", out var p)
                ? p.Clone()
                : JsonSerializer.SerializeToElement(new { });
            if (parameters.ValueKind != JsonValueKind.Object)
            {
                throw new OfficeArgumentException("JSON-RPC params must be an object");
            }

            rpcMethod = Catalog.FindRpcMethod(method);
            if (rpcMethod is not null)
            {
                Catalog.ValidateRpcParams(rpcMethod, parameters);
            }

            if (method == "office.command.execute")
            {
                Catalog.ValidateCommandParams(parameters);
                commandPolicy = CommandPolicy.Evaluate(
                    parameters,
                    Catalog.SecurityPolicy,
                    _workspaceRoot);
            }

            var stopwatch = Stopwatch.StartNew();
            object result = Execute(method, parameters, commandPolicy);
            stopwatch.Stop();

            if (rpcMethod is not null)
            {
                Catalog.ValidateRpcResult(
                    rpcMethod,
                    JsonSerializer.SerializeToElement(result));
            }

            object finalResult = method == "office.command.execute"
                ? AddAudit(
                    result,
                    commandPolicy ?? throw new InvalidOperationException("command policy was not evaluated"),
                    stopwatch.ElapsedMilliseconds)
                : result;
            return JsonSerializer.Serialize(new { jsonrpc = "2.0", id, result = finalResult });
        }
        catch (OfficeComException ex)
        {
            if (ex.Code == OfficeErrorCode.OfficeModalDialog)
            {
                _notifications.PublishEvent(
                    "office.modal.detected",
                    $"rpc:{id.GetRawText()}",
                    new { error = ex.Message });
            }
            if (ex.Code == OfficeErrorCode.OfficeUserConfirmationRequired)
            {
                _notifications.PublishEvent(
                    "office.security.prompt",
                    $"rpc:{id.GetRawText()}",
                    new { error = ex.Message });
            }
            return Error(id, ex.Code, ex.Message, ex.Indeterminate);
        }
        catch (OfficeArgumentException ex)
        {
            return Error(id, OfficeErrorCode.OfficeInvalidRequest, ex.Message);
        }
        catch (JsonException ex)
        {
            return Error(id, OfficeErrorCode.OfficeInvalidRequest, ex.Message);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[office-host:{_app}] unhandled: {ex}");
            return Error(id, OfficeErrorCode.OfficeBackendUnavailable, ex.Message);
        }
    }

    internal static string Error(
        JsonElement id,
        OfficeErrorCode code,
        string message,
        bool indeterminate = false)
    {
        var error = new JsonObject
        {
            ["code"] = code.ToWireName(),
            ["message"] = message,
        };
        if (indeterminate)
        {
            error["data"] = new JsonObject { ["indeterminate"] = true };
        }
        return JsonSerializer.Serialize(new { jsonrpc = "2.0", id, error });
    }

    // ------------------------------------------------------------- methods

    private object Execute(
        string method,
        JsonElement parameters,
        CommandPolicy? commandPolicy) => method switch
        {
            "office.host.ping" => Ping(),
            "office.host.handshake" => Handshake(parameters),
            "office.host.shutdown" => Shutdown(),
            "office.job.get" => JobGet(parameters),
            "office.job.cancel" => JobCancel(parameters),
            "office.command.execute" => ExecuteCommand(
                parameters,
                commandPolicy ?? throw new InvalidOperationException("command policy is required")),
            _ => throw new OfficeArgumentException($"unknown method: {method}"),
        };

    /// <summary>
    /// Graceful sidecar stop: replies, then the pipe loop exits after this
    /// connection and Dispose quits the COM app — no orphaned Office
    /// processes when the gateway tears a sidecar down (proposal §8.3).
    /// </summary>
    private object Shutdown()
    {
        ShutdownRequested = true;
        _notifications.PublishEvent(
            "office.application.stopped",
            "host:shutdown",
            new { reason = "office.host.shutdown" });
        return new { ok = true };
    }

    private object Ping()
    {
        bool busy = _jobs.IsBusy;
        bool modal = _com?.OfficeProcessId is int pid
            && ModalDialogDetector.FindModalDialogTitle(pid) is not null;
        string attachState = _com is null
            ? "unavailable"
            : _comAttached
                ? "attached"
                : _comAttachAttempted
                    ? "failed"
                    : "unknown";
        return new
        {
            protocol_version = HostBuildInfo.ProtocolVersion,
            host_id = HostId(),
            state = busy ? "busy" : attachState == "failed" ? "degraded" : "ready",
            app = _app,
            pid = Environment.ProcessId,
            office_version = (string?)null,
            open_documents = Array.Empty<string>(),
            busy,
            modal,
            protected_view = false,
            com_attached = _comAttached,
            com_attach_state = attachState,
        };
    }

    private object Handshake(JsonElement parameters)
    {
        string? requestedApp = parameters.TryGetProperty("requested_app", out var appElement)
            ? appElement.GetString()
            : null;
        if (requestedApp is null || !requestedApp.Equals(_app, StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeArgumentException(
                $"requested_app '{requestedApp}' does not match this sidecar's app '{_app}'");
        }
        EnsureComAttached();
        if (!_applicationStartedEventPublished)
        {
            _applicationStartedEventPublished = true;
            _notifications.PublishEvent(
                "office.application.started",
                "host:handshake",
                new { com_attach_state = _comAttached ? "attached" : "failed" });
        }
        return new
        {
            protocol_version = HostBuildInfo.ProtocolVersion,
            host_id = HostId(),
            capability_manifest = Manifest(),
        };
    }

    private object JobGet(JsonElement parameters)
    {
        string jobId = parameters.GetProperty("job_id").GetString()
            ?? throw new OfficeArgumentException("params.job_id is required");
        return _jobs.Get(jobId);
    }

    private object JobCancel(JsonElement parameters)
    {
        string jobId = parameters.GetProperty("job_id").GetString()
            ?? throw new OfficeArgumentException("params.job_id is required");
        return _jobs.Cancel(jobId);
    }

    private object ExecuteCommand(JsonElement parameters, CommandPolicy policy)
    {
        string capability = parameters.TryGetProperty("capability", out var c)
            ? c.GetString() ?? ""
            : throw new OfficeArgumentException("capability is required");
        var input = parameters.TryGetProperty("input", out var i) ? i.Clone() : default;
        CapabilityDefinition definition = Catalog.FindCapability(capability)
            ?? throw new OfficeComException(
                OfficeErrorCode.OfficeCapabilityUnsupported,
                $"capability '{capability}' is not in the office-rpc catalog");
        Catalog.ValidateInput(definition, input);
        if (parameters.TryGetProperty("document", out JsonElement document)
            && document.ValueKind != JsonValueKind.Null)
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeCapabilityUnsupported,
                "document.expected_revision is not implemented; guarded requests are refused instead of silently overwriting");
        }
        policy.ValidateWorkspace(definition, input);
        if (definition.HandlerId == CapabilityHandler.SlideRender)
        {
            RefuseExistingRenderOutputs(input);
        }
        if (definition.HandlerId == CapabilityHandler.BatchReplaceText
            && input.TryGetProperty("dry_run", out JsonElement dryRun)
            && dryRun.ValueKind == JsonValueKind.False)
        {
            policy.RequireConfirmation(parameters, "overwrite_original");
        }
        if (definition.HandlerId == CapabilityHandler.BatchConvert
            && input.TryGetProperty("overwrite", out JsonElement overwrite)
            && overwrite.ValueKind == JsonValueKind.String
            && overwrite.GetString() == "overwrite")
        {
            policy.RequireConfirmation(parameters, "overwrite_original");
        }
        if (definition.HandlerId is CapabilityHandler.BatchConvert
            or CapabilityHandler.BatchReplaceText)
        {
            ValidateBatchSubmission(definition.HandlerId, input);
            object submission = SubmitBatch(definition, input, policy);
            Catalog.ValidateOutput(
                definition,
                JsonSerializer.SerializeToElement(submission));
            return submission;
        }

        object result = ExecuteCapability(definition, input, job: null);
        Catalog.ValidateOutput(definition, JsonSerializer.SerializeToElement(result));
        PublishCapabilityEvents(definition.HandlerId, input, result);
        return result;
    }

    private static void ValidateBatchSubmission(
        CapabilityHandler handler,
        JsonElement input)
    {
        if (handler != CapabilityHandler.BatchConvert)
        {
            return;
        }
        string targetFormat = input.TryGetProperty("target_format", out JsonElement target)
            ? target.GetString() ?? "pdf"
            : "pdf";
        if (!targetFormat.Equals("pdf", StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeCapabilityUnsupported,
                $"target_format '{targetFormat}' is not supported (pdf only)");
        }
        string backend = input.TryGetProperty("backend", out JsonElement requestedBackend)
            ? requestedBackend.GetString() ?? "auto"
            : "auto";
        if (backend is not ("auto" or "desktop_com"))
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeCapabilityUnsupported,
                $"batch.convert backend '{backend}' is not implemented by the desktop sidecar");
        }
    }

    private object SubmitBatch(
        CapabilityDefinition definition,
        JsonElement input,
        CommandPolicy policy)
    {
        int total = input.TryGetProperty("inputs", out JsonElement inputs)
            && inputs.ValueKind == JsonValueKind.Array
            ? inputs.GetArrayLength()
            : 0;
        JobSnapshot job = _jobs.Submit(definition.Name, total, context =>
        {
            var stopwatch = Stopwatch.StartNew();
            object result = ExecuteCapability(definition, input, context);
            stopwatch.Stop();
            Catalog.ValidateOutput(definition, JsonSerializer.SerializeToElement(result));
            return AddAudit(result, policy, stopwatch.ElapsedMilliseconds);
        });
        return new
        {
            operation_id = job.JobId,
            job_id = job.JobId,
            phase = job.Phase,
            changed = new
            {
                job_id = job.JobId,
                capability = definition.Name,
                phase = job.Phase,
            },
            warnings = Array.Empty<string>(),
            artefacts = Array.Empty<object>(),
            validation = new { accepted = true },
            backend = "job",
            indeterminate = false,
        };
    }

    private object ExecuteCapability(
        CapabilityDefinition definition,
        JsonElement input,
        InMemoryJobTracker.JobExecutionContext? job) => definition.HandlerId switch
        {
            CapabilityHandler.DeckCompile => Compile(input),
            CapabilityHandler.DocumentInspect => InspectDocument(input),
            CapabilityHandler.BatchConvert => BatchConvert(input, job),
            CapabilityHandler.BatchReplaceText => BatchReplaceText(input, job),
            CapabilityHandler.SlideRender => RenderSlides(input),
            _ => throw new InvalidOperationException(
                $"catalog handler '{definition.Handler}' is not routed"),
        };

    private void PublishCapabilityEvents(
        CapabilityHandler handler,
        JsonElement input,
        object result)
    {
        string correlationId = JsonSerializer.SerializeToNode(result)?["operation_id"]?
            .GetValue<string>() ?? $"operation:{Guid.NewGuid():N}";
        if (handler == CapabilityHandler.DeckCompile)
        {
            _notifications.PublishEvent(
                "office.document.saved",
                correlationId,
                new { path = input.GetProperty("output").GetString() });
        }
        if (handler == CapabilityHandler.DocumentInspect)
        {
            string? path = input.GetProperty("path").GetString();
            _notifications.PublishEvent(
                "office.document.opened",
                correlationId,
                new { path });
            _notifications.PublishEvent(
                "office.document.before_close",
                correlationId,
                new { path });
        }
    }

    // ---------------------------------------------------------- capabilities

    /// <summary>deck.compile — Deck IR → PPTX via the Open XML worker (no Office).</summary>
    private object Compile(JsonElement input)
    {
        string ir = input.TryGetProperty("ir", out var irElement) && irElement.ValueKind == JsonValueKind.String
            ? irElement.GetString()!
            : throw new OfficeArgumentException("input.ir (inline JSON or a path to an IR file) is required");
        string output = input.TryGetProperty("output", out var outElement) && outElement.ValueKind == JsonValueKind.String
            ? outElement.GetString()!
            : throw new OfficeArgumentException("input.output is required");
        output = Path.GetFullPath(output);

        string templateUri = InputTemplateUri(input, ir) ?? TemplateRegistry.DefaultUri;
        TemplateEntry template = _templates.Resolve(templateUri)
            ?? throw new OfficeComException(
                OfficeErrorCode.OfficeCapabilityUnsupported,
                $"brand template '{templateUri}' is not materialized; available: {string.Join(", ", _templates.AllUris)}");
        var warnings = new List<string>();
        warnings.AddRange(CheckLayoutsAgainstTemplate(ir, template));

        if (File.Exists(output))
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeAccessDenied,
                "deck.compile refuses an existing output because the capability has no overwrite mode");
        }

        string operationId = Guid.NewGuid().ToString("N");
        string outputDirectory = Path.GetDirectoryName(output)!;
        Directory.CreateDirectory(outputDirectory);
        string stagedOutput = Path.Combine(
            outputDirectory,
            $".{Path.GetFileNameWithoutExtension(output)}.dcc-stage-{operationId}.pptx");
        PptxInspector.DeckInfo info;
        try
        {
            if (ir.TrimStart().StartsWith('{'))
            {
                // Keep transient IR inside the bound workspace as well.
                string irTemp = Path.Combine(outputDirectory, $".dcc-ir-{operationId}.json");
                File.WriteAllText(irTemp, ir);
                try
                {
                    PptxWriter.CompileDeck(irTemp, stagedOutput, template.Package);
                }
                finally
                {
                    File.Delete(irTemp);
                }
            }
            else
            {
                PptxWriter.CompileDeck(ir, stagedOutput, template.Package);
            }
            info = PptxInspector.Inspect(stagedOutput);
            File.Move(stagedOutput, output, overwrite: false);
        }
        finally
        {
            File.Delete(stagedOutput);
        }
        return new
        {
            operation_id = operationId,
            changed = new { files = 1, slides = info.SlideCount },
            warnings,
            artefacts = new[] { Artifact(output, "pptx") },
            validation = new { output_openable = true, non_empty = true, slide_count_reasonable = info.SlideCount >= 1 },
            backend = "openxml",
            indeterminate = false,
        };
    }

    /// <summary>
    /// Best-effort semantic-layout check against the built-in brand package:
    /// unknown layouts still compile (the worker falls back to bullets) but
    /// the result carries a warning so agents notice the substitution.
    /// </summary>
    private static string? InputTemplateUri(JsonElement input, string ir)
    {
        if (input.TryGetProperty("template", out JsonElement requested)
            && requested.ValueKind == JsonValueKind.String)
        {
            return requested.GetString();
        }
        string json = ir.TrimStart().StartsWith('{') ? ir : File.ReadAllText(ir);
        using JsonDocument document = JsonDocument.Parse(json);
        return document.RootElement.TryGetProperty("template", out JsonElement template)
            && template.ValueKind == JsonValueKind.Object
            && template.TryGetProperty("uri", out JsonElement uri)
            && uri.ValueKind == JsonValueKind.String
            ? uri.GetString()
            : null;
    }

    private static List<string> CheckLayoutsAgainstTemplate(string ir, TemplateEntry template)
    {
        var warnings = new List<string>();
        try
        {
            string json = ir.TrimStart().StartsWith('{') ? ir : File.ReadAllText(ir);
            using var document = JsonDocument.Parse(json);
            if (document.RootElement.TryGetProperty("document", out var doc)
                && doc.TryGetProperty("slides", out var slides)
                && slides.ValueKind == JsonValueKind.Array)
            {
                var known = template.Package.Layouts.ToHashSet(StringComparer.Ordinal);
                foreach (var slide in slides.EnumerateArray())
                {
                    var layout = slide.TryGetProperty("semantic_layout", out var sl)
                        ? sl.GetString() ?? ""
                        : "";
                    if (layout.Length > 0 && !known.Contains(layout))
                    {
                        warnings.Add(
                            $"semantic_layout '{layout}' is not in {template.Package.Uri}; compiled as 'bullets'");
                    }
                }
            }
        }
        catch
        {
            // Layout checking is advisory; malformed IR fails in the worker.
        }
        return warnings;
    }

    private object InspectDocument(JsonElement input)
    {
        string path = input.TryGetProperty("path", out var pathElement) && pathElement.ValueKind == JsonValueKind.String
            ? pathElement.GetString()!
            : throw new OfficeArgumentException("input.path is required");
        string backend = input.TryGetProperty("backend", out var b) ? b.GetString() ?? "auto" : "auto";
        bool useCom = backend switch
        {
            "auto" => ComAttached,
            "desktop_com" => true,
            "openxml" => false,
            _ => throw new OfficeArgumentException($"unknown backend '{backend}' (auto | desktop_com | openxml)"),
        };
        object summary;
        string usedBackend;
        if (useCom)
        {
            var outcome = Com().Inspect(Path.GetFullPath(path));
            var flat = new JsonObject
            {
                ["path"] = outcome.Path,
                ["kind"] = outcome.Kind,
            };
            foreach (var pair in outcome.Summary)
            {
                flat[pair.Key] = pair.Value?.DeepClone();
            }
            summary = flat;
            usedBackend = outcome.Backend;
        }
        else
        {
            var info = PptxInspector.Inspect(path); // throws on non-pptx
            summary = new
            {
                slide_count = info.SlideCount,
                title = info.Title,
                slides = info.Slides.Select(s => new
                {
                    index = s.Index,
                    shapes = s.ShapeCount,
                    pictures = s.Pictures,
                    pictures_without_alt = s.PicturesWithoutAlt,
                    has_notes = s.HasNotes,
                }).ToArray(),
            };
            usedBackend = "openxml";
        }
        return new
        {
            operation_id = Guid.NewGuid().ToString("N"),
            changed = new { summary },
            warnings = Array.Empty<string>(),
            artefacts = Array.Empty<object>(),
            validation = new { },
            backend = usedBackend,
            indeterminate = false,
        };
    }

    private object BatchConvert(
        JsonElement input,
        InMemoryJobTracker.JobExecutionContext? job)
    {
        string targetFormat = input.TryGetProperty("target_format", out var tf) ? tf.GetString() ?? "pdf" : "pdf";
        if (!targetFormat.Equals("pdf", StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeComException(OfficeErrorCode.OfficeCapabilityUnsupported,
                $"target_format '{targetFormat}' is not supported (pdf only)");
        }
        string requestedBackend = input.TryGetProperty("backend", out var backendElement)
            ? backendElement.GetString() ?? "auto"
            : "auto";
        if (requestedBackend is not ("auto" or "desktop_com"))
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeCapabilityUnsupported,
                $"batch.convert backend '{requestedBackend}' is not implemented by the desktop sidecar");
        }
        string outputDir = input.TryGetProperty("output_directory", out var od) && od.ValueKind == JsonValueKind.String
            ? od.GetString()!
            : throw new OfficeArgumentException("input.output_directory is required");
        string overwrite = input.TryGetProperty("overwrite", out var overwriteElement)
            ? overwriteElement.GetString() ?? "versioned"
            : "versioned";
        string[] requestedValidation = input.TryGetProperty("validation", out var validationElement)
            ? validationElement.EnumerateArray().Select(item => item.GetString() ?? "").ToArray()
            : ["output_openable", "non_empty", "page_count_reasonable"];
        outputDir = Path.GetFullPath(outputDir);
        Directory.CreateDirectory(outputDir);

        var paths = ResolveInputs(input);
        if (paths.Length == 0)
        {
            throw new OfficeArgumentException("input.inputs resolved to no files");
        }
        job?.SetTotal(paths.Length);

        string expected = ExpectedExtension(_app);
        string operationId = Guid.NewGuid().ToString("N");
        var reservedOutputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var outputPlan = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string path in paths.Where(path =>
            Path.GetExtension(path).Equals(expected, StringComparison.OrdinalIgnoreCase)))
        {
            string outputPath = ReserveOutputPathForMode(
                outputDir,
                Path.GetFileNameWithoutExtension(path) + ".pdf",
                overwrite,
                reservedOutputs);
            outputPlan.Add(path, outputPath);
        }
        WorkspaceGuard.ValidatePaths(outputPlan.Values, _workspaceRoot);

        var items = new JsonArray();
        var artefacts = new JsonArray();
        if (overwrite == "overwrite")
        {
            foreach (string outputPath in outputPlan.Values.Where(File.Exists))
            {
                artefacts.Add(CreateCheckpoint(outputPath, operationId));
            }
        }
        var warnings = new JsonArray();
        var successful = new List<FileConvertOutcome>();
        int processed = 0;
        int succeeded = 0;
        bool indeterminate = false;
        OfficeComBackend? com = null;
        foreach (var path in paths)
        {
            if (job?.StopBeforeNextItem() == true)
            {
                break;
            }
            var ext = Path.GetExtension(path);
            if (!ext.Equals(expected, StringComparison.OrdinalIgnoreCase))
            {
                items.Add(new JsonObject
                {
                    ["input_path"] = path,
                    ["ok"] = false,
                    ["error_code"] = OfficeErrorCode.OfficeCapabilityUnsupported.ToWireName(),
                    ["error"] = $"{_app} sidecar handles {expected} only; route '{path}' to its own sidecar",
                });
                processed++;
                job?.Report("converting", processed);
                continue;
            }
            string outputPath = outputPlan[path];
            if (overwrite == "overwrite" && File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            com ??= Com();
            var outcome = com.ConvertToPdf(path, outputPath);
            indeterminate |= outcome.Indeterminate;
            if (outcome.Ok)
            {
                succeeded++;
                successful.Add(outcome);
                artefacts.Add(Artifact(outputPath, "pdf"));
                _notifications.PublishEvent(
                    "office.document.saved",
                    job?.JobId ?? operationId,
                    new { input_path = path, output_path = outputPath });
            }
            else if (outcome.ErrorCode is not null)
            {
                warnings.Add($"{path}: {outcome.Error}");
            }
            items.Add(JsonSerializer.SerializeToNode(outcome)!);
            processed++;
            job?.Report("converting", processed);
        }
        var validation = new JsonObject();
        foreach (string check in requestedValidation)
        {
            validation[check] = check switch
            {
                "output_openable" => successful.Count > 0
                    && successful.All(outcome =>
                        outcome.OutputPath is not null && File.Exists(outcome.OutputPath)),
                "non_empty" => successful.Count > 0
                    && successful.All(outcome => outcome.Bytes > 0),
                "page_count_reasonable" => successful.Count > 0
                    && successful.All(outcome => outcome.PageCount > 0),
                _ => throw new InvalidOperationException(
                    $"schema allowed unknown validation check '{check}'"),
            };
        }
        return new
        {
            operation_id = operationId,
            changed = new
            {
                files = paths.Length,
                processed,
                succeeded,
                failed = processed - succeeded,
                cancelled = job?.CancellationObserved == true,
                target_format = targetFormat,
                items,
            },
            warnings,
            artefacts,
            validation,
            backend = "desktop_com",
            indeterminate,
        };
    }

    private object BatchReplaceText(
        JsonElement input,
        InMemoryJobTracker.JobExecutionContext? job)
    {
        string operationId = Guid.NewGuid().ToString("N");
        var rules = new List<ReplaceRuleInput>();
        if (input.TryGetProperty("rules", out var rulesElement) && rulesElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var rule in rulesElement.EnumerateArray())
            {
                rules.Add(new ReplaceRuleInput
                {
                    Find = rule.TryGetProperty("find", out var f) ? f.GetString() ?? "" : "",
                    Replace = rule.TryGetProperty("replace", out var r) ? r.GetString() ?? "" : "",
                    Match = rule.TryGetProperty("match", out var m) ? m.GetString() ?? "literal" : "literal",
                });
            }
        }
        if (rules.Count == 0)
        {
            throw new OfficeArgumentException("input.rules must contain at least one rule");
        }
        var scope = new List<string>();
        if (input.TryGetProperty("scope", out var scopeElement) && scopeElement.ValueKind == JsonValueKind.Array)
        {
            scope.AddRange(scopeElement.EnumerateArray().Select(s => s.GetString() ?? ""));
        }
        bool dryRun = !input.TryGetProperty("dry_run", out var dryElement) || dryElement.ValueKind != JsonValueKind.False;

        var paths = ResolveInputs(input);
        if (paths.Length == 0)
        {
            throw new OfficeArgumentException("input.inputs resolved to no files");
        }
        job?.SetTotal(paths.Length);

        string expected = ExpectedExtension(_app);
        var items = new JsonArray();
        var artefacts = new JsonArray();
        if (!dryRun)
        {
            foreach (string path in paths.Where(path =>
                Path.GetExtension(path).Equals(expected, StringComparison.OrdinalIgnoreCase)))
            {
                artefacts.Add(CreateCheckpoint(path, operationId));
            }
        }
        int totalMatched = 0;
        int totalReplaced = 0;
        int processed = 0;
        int succeeded = 0;
        bool indeterminate = false;
        OfficeComBackend? com = null;
        foreach (var path in paths)
        {
            if (job?.StopBeforeNextItem() == true)
            {
                break;
            }
            var ext = Path.GetExtension(path);
            if (!ext.Equals(expected, StringComparison.OrdinalIgnoreCase))
            {
                items.Add(new JsonObject
                {
                    ["path"] = path,
                    ["error"] = $"{_app} sidecar handles {expected} only",
                });
                processed++;
                job?.Report("replacing", processed);
                continue;
            }
            com ??= Com();
            var outcome = com.ReplaceText(path, rules, scope, dryRun);
            indeterminate |= outcome.Indeterminate;
            totalMatched += outcome.TotalMatched;
            totalReplaced += outcome.TotalReplaced;
            items.Add(JsonSerializer.SerializeToNode(outcome)!);
            bool itemSucceeded = !outcome.Warnings.Any(warning =>
                warning.StartsWith("replace failed:", StringComparison.Ordinal));
            if (itemSucceeded)
            {
                succeeded++;
                if (!dryRun)
                {
                    _notifications.PublishEvent(
                        "office.document.changed",
                        job?.JobId ?? operationId,
                        new { path, total_replaced = outcome.TotalReplaced });
                    _notifications.PublishEvent(
                        "office.document.saved",
                        job?.JobId ?? operationId,
                        new { path });
                }
            }
            processed++;
            job?.Report("replacing", processed);
        }
        return new
        {
            operation_id = operationId,
            changed = new
            {
                files = paths.Length,
                processed,
                succeeded,
                failed = processed - succeeded,
                cancelled = job?.CancellationObserved == true,
                dry_run = dryRun,
                total_matched = totalMatched,
                total_replaced = totalReplaced,
                items,
            },
            warnings = Array.Empty<string>(),
            artefacts,
            validation = new { },
            backend = "desktop_com",
            indeterminate,
        };
    }

    private object RenderSlides(JsonElement input)
    {
        var com = Com();
        if (!_app.Equals("powerpoint", StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeComException(OfficeErrorCode.OfficeCapabilityUnsupported,
                "slide.render is a PowerPoint capability");
        }
        string path = input.TryGetProperty("path", out var pathElement) && pathElement.ValueKind == JsonValueKind.String
            ? pathElement.GetString()!
            : throw new OfficeArgumentException("input.path is required");
        string outputDir = input.TryGetProperty("output_directory", out var od) && od.ValueKind == JsonValueKind.String
            ? od.GetString()!
            : throw new OfficeArgumentException("input.output_directory is required");
        int width = input.TryGetProperty("width", out var w) ? w.GetInt32() : 1280;
        int height = input.TryGetProperty("height", out var h) ? h.GetInt32() : 720;

        // COM resolves relative paths against the Office process working
        // directory (usually System32), so relative input is guaranteed to
        // fail with 0x80070003 — resolve both to absolute paths before COM.
        path = Path.GetFullPath(path);
        outputDir = Path.GetFullPath(outputDir);

        var previews = com.ExportSlidePreviews(path, outputDir, width, height)
            ?? throw new OfficeComException(OfficeErrorCode.OfficeBackendUnavailable,
                $"{_app} backend cannot render slide previews");
        var artefacts = new JsonArray();
        int ok = 0;
        var overflow = new JsonArray();
        foreach (var preview in previews)
        {
            if (preview.Ok && preview.Path is not null)
            {
                ok++;
                artefacts.Add(Artifact(preview.Path, "png"));
            }
            foreach (var shape in preview.Overflow)
            {
                overflow.Add(new JsonObject
                {
                    ["slide"] = preview.SlideIndex,
                    ["shape_id"] = shape.ShapeId,
                    ["name"] = shape.Name,
                    ["kind"] = shape.Kind,
                    ["left"] = shape.Left,
                    ["top"] = shape.Top,
                    ["right"] = shape.Right,
                    ["bottom"] = shape.Bottom,
                });
            }
        }
        return new
        {
            operation_id = Guid.NewGuid().ToString("N"),
            changed = new
            {
                previews = previews.Count,
                ok,
                failed = previews.Count - ok,
                width,
                height,
                overflow,
            },
            warnings = Array.Empty<string>(),
            artefacts,
            validation = new { },
            backend = "desktop_com",
            indeterminate = false,
        };
    }

    // ------------------------------------------------------------- helpers

    private OfficeComBackend Com()
    {
        EnsureComAttached();
        if (_com is null || !ComAttached)
        {
            throw new OfficeComException(OfficeErrorCode.OfficeBackendUnavailable,
                $"{_app} desktop COM backend is not available in this session");
        }
        return _com;
    }

    private void EnsureComAttached()
    {
        if (_com is null)
        {
            return;
        }
        lock (_attachLock)
        {
            if (_comAttached)
            {
                return;
            }
            _comAttachAttempted = true;
            _comAttached = _com.TryAttach(AttachBudget);
        }
    }

    private object Manifest()
    {
        IReadOnlyDictionary<string, string> capabilities =
            Catalog.ManifestCapabilities(_app, _comAttached);
        IReadOnlyList<string> modes = Catalog.ManifestExecutionModes(_app, _comAttached);
        object? application = null;
        if (_comAttached)
        {
            var info = _com!.GetApplicationInfo();
            application = new
            {
                name = info.Name,
                version = info.Version,
                bitness = info.Bitness,
                language = info.Language,
            };
        }
        return new
        {
            provider = Catalog.Provider,
            provider_version = HostBuildInfo.Version,
            protocol_version = Catalog.ProtocolVersion,
            application,
            execution_modes = modes,
            capabilities,
            template_packages = _templates.Capabilities,
            limits = new { max_parallel_writes = 1, requires_user_session = true },
        };
    }

    private string HostId() =>
        $"office-host:{_app}:session-{System.Diagnostics.Process.GetCurrentProcess().SessionId}";

    private static string ExpectedExtension(string app) => app.ToLowerInvariant() switch
    {
        "powerpoint" => ".pptx",
        "word" => ".docx",
        "excel" => ".xlsx",
        _ => throw new OfficeComException(OfficeErrorCode.OfficeCapabilityUnsupported,
            $"no desktop COM backend for app '{app}'"),
    };

    private static void RefuseExistingRenderOutputs(JsonElement input)
    {
        if (input.TryGetProperty("output_directory", out JsonElement directory)
            && directory.ValueKind == JsonValueKind.String
            && Directory.Exists(Path.GetFullPath(directory.GetString()!))
            && Directory.EnumerateFiles(
                Path.GetFullPath(directory.GetString()!),
                "slide-*.png",
                SearchOption.TopDirectoryOnly).Any())
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeAccessDenied,
                "slide.render refuses existing slide-*.png outputs; choose an empty output directory");
        }
    }

    internal static string OutputPathForMode(string directory, string fileName, string overwrite)
    {
        var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        return ReserveOutputPathForMode(directory, fileName, overwrite, reserved);
    }

    private static string ReserveOutputPathForMode(
        string directory,
        string fileName,
        string overwrite,
        ISet<string> reserved)
    {
        string candidate = Path.GetFullPath(Path.Combine(directory, fileName));
        if (overwrite == "overwrite")
        {
            if (!reserved.Add(candidate))
            {
                throw new OfficeArgumentException(
                    $"multiple inputs resolve to the same overwrite output: {candidate}");
            }
            return candidate;
        }
        if (overwrite == "fail")
        {
            if (File.Exists(candidate))
            {
                throw new OfficeArgumentException(
                    $"output already exists and overwrite is 'fail': {candidate}");
            }
            if (!reserved.Add(candidate))
            {
                throw new OfficeArgumentException(
                    $"multiple inputs resolve to the same output: {candidate}");
            }
            return candidate;
        }
        if (overwrite != "versioned")
        {
            throw new OfficeArgumentException($"unknown overwrite mode '{overwrite}'");
        }
        if (!File.Exists(candidate) && reserved.Add(candidate))
        {
            return candidate;
        }
        string stem = Path.GetFileNameWithoutExtension(fileName);
        string ext = Path.GetExtension(fileName);
        for (int n = 2; ; n++)
        {
            candidate = Path.GetFullPath(Path.Combine(directory, $"{stem}.v{n}{ext}"));
            if (!File.Exists(candidate) && reserved.Add(candidate))
            {
                return candidate;
            }
        }
    }

    private static JsonObject Artifact(string path, string kind)
    {
        string full = Path.GetFullPath(path);
        return new JsonObject
        {
            ["artifact_id"] = Guid.NewGuid().ToString("N"),
            ["kind"] = kind,
            ["path"] = full,
            ["sha256"] = Sha256(full),
        };
    }

    /// <summary>Creates the mandatory byte-exact pre-image before an in-place write.</summary>
    internal static JsonObject CreateCheckpoint(string path, string operationId)
    {
        string source = Path.GetFullPath(path);
        string safeOperation = Regex.Replace(operationId, "[^A-Za-z0-9_-]", "");
        if (safeOperation.Length == 0)
        {
            throw new OfficeArgumentException("operation id cannot name a checkpoint");
        }
        string checkpoint = Path.Combine(
            Path.GetDirectoryName(source)!,
            $"{Path.GetFileNameWithoutExtension(source)}.dcc-checkpoint-{safeOperation}{Path.GetExtension(source)}");
        try
        {
            File.Copy(source, checkpoint, overwrite: false);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeAccessDenied,
                $"checkpoint creation failed before writing '{source}': {ex.Message}",
                ex);
        }
        return Artifact(checkpoint, "checkpoint");
    }

    private static string Sha256(string path)
    {
        IOException? lastError = null;
        for (int attempt = 0; attempt < 10; attempt++)
        {
            try
            {
                using var stream = new FileStream(
                    path,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete);
                return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
            }
            catch (IOException ex)
            {
                lastError = ex;
                Thread.Sleep(25);
            }
        }
        throw new IOException($"could not hash artifact after bounded retries: {path}", lastError);
    }

    /// <summary>
    /// Resolves input.inputs: plain paths plus simple wildcard patterns
    /// (* and ?, with **/ for recursion) — proposal §15.1 glob surface.
    /// </summary>
    private static string[] ResolveInputs(JsonElement input)
    {
        var paths = new List<string>();
        if (input.TryGetProperty("inputs", out var inputsElement) && inputsElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in inputsElement.EnumerateArray())
            {
                var spec = item.GetString();
                if (string.IsNullOrWhiteSpace(spec))
                {
                    continue;
                }
                if (spec.Contains('*') || spec.Contains('?'))
                {
                    paths.AddRange(ExpandGlob(spec));
                }
                else
                {
                    paths.Add(Path.GetFullPath(spec));
                }
            }
        }
        return paths.Where(File.Exists).Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static IEnumerable<string> ExpandGlob(string pattern)
    {
        string norm = pattern.Replace('/', '\\');
        string regexPattern = "^" + Regex.Escape(norm)
            .Replace(@"\*\*\\", @"(?:.*\\)?")
            .Replace(@"\*\*", ".*")
            .Replace(@"\*", @"[^\\]*")
            .Replace(@"\?", @"[^\\]") + "$";
        var regex = new Regex(regexPattern, RegexOptions.IgnoreCase);

        int wildcardAt = norm.IndexOfAny(new[] { '*', '?' });
        if (wildcardAt < 0)
        {
            return new[] { norm };
        }
        string literalPrefix = norm[..wildcardAt];
        string startDir = Directory.Exists(literalPrefix)
            ? literalPrefix
            : Path.GetDirectoryName(literalPrefix) ?? Directory.GetCurrentDirectory();
        if (startDir.Length == 0)
        {
            startDir = Directory.GetCurrentDirectory();
        }
        return Directory.EnumerateFiles(startDir, "*", SearchOption.AllDirectories)
            .Where(f => regex.IsMatch(f))
            .ToArray();
    }

    public void Dispose()
    {
        _jobs.Dispose();
        _com?.Dispose();
        _sta.Dispose();
    }
}

/// <summary>Invalid-request error mapped to OFFICE_INVALID_REQUEST on the wire.</summary>
public sealed class OfficeArgumentException : Exception
{
    public OfficeArgumentException(string message)
        : base(message)
    {
    }
}
