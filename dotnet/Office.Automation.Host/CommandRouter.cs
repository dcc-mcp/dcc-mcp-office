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
    private const string ProtocolVersion = "office-rpc/1";
    private const string ProviderVersion = "0.1.1";

    private static readonly TimeSpan AttachBudget = TimeSpan.FromSeconds(60);

    private readonly string _app;
    private readonly StaDispatcher _sta = new();
    private readonly OfficeComBackend? _com;
    private readonly object _attachLock = new();
    private bool _comAttached;

    public CommandRouter(string app)
    {
        _app = app;
        _com = ComBackendFactory.IsSupported(app)
            ? ComBackendFactory.Create(app, _sta)
            : null;
    }

    public string App => _app;

    public bool ComAttached => _comAttached;

    /// <summary>Processes one JSON-RPC request line and returns the response line.</summary>
    public string Dispatch(string requestJson)
    {
        JsonElement id = default;
        string method;
        try
        {
            using var request = JsonDocument.Parse(requestJson);
            var root = request.RootElement;
            id = root.TryGetProperty("id", out var idElement) ? idElement.Clone() : default;
            method = root.GetProperty("method").GetString() ?? "";
            var parameters = root.TryGetProperty("params", out var p) ? p.Clone() : default;
            var result = Execute(method, parameters);
            return JsonSerializer.Serialize(new { jsonrpc = "2.0", id, result });
        }
        catch (OfficeComException ex)
        {
            return Error(id, ex.Code.ToWireName(), ex.Message);
        }
        catch (OfficeArgumentException ex)
        {
            return Error(id, "OFFICE_INVALID_REQUEST", ex.Message);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[office-host:{_app}] unhandled: {ex}");
            return Error(id, "OFFICE_BACKEND_UNAVAILABLE", ex.Message);
        }
    }

    private static string Error(JsonElement id, string code, string message) =>
        JsonSerializer.Serialize(new { jsonrpc = "2.0", id, error = new { code, message } });

    // ------------------------------------------------------------- methods

    private object Execute(string method, JsonElement parameters) => method switch
    {
        "office.host.ping" => Ping(),
        "office.host.handshake" => Handshake(parameters),
        "office.command.execute" => ExecuteCommand(parameters),
        _ => throw new OfficeArgumentException($"unknown method: {method}"),
    };

    private object Ping()
    {
        EnsureComAttached();
        return new
        {
            app = _app,
            protocol_version = ProtocolVersion,
            host_id = HostId(),
            com_attached = _comAttached,
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
        return new
        {
            protocol_version = ProtocolVersion,
            host_id = HostId(),
            capability_manifest = Manifest(),
        };
    }

    private object ExecuteCommand(JsonElement parameters)
    {
        string capability = parameters.TryGetProperty("capability", out var c)
            ? c.GetString() ?? ""
            : throw new OfficeArgumentException("capability is required");
        var input = parameters.TryGetProperty("input", out var i) ? i.Clone() : default;
        return capability switch
        {
            "deck.compile" => Compile(input),
            "document.inspect" => InspectDocument(input),
            "batch.convert" => BatchConvert(input),
            "batch.replace_text" => BatchReplaceText(input),
            "slide.render" => RenderSlides(input),
            _ => throw new OfficeComException(OfficeErrorCode.OfficeCapabilityUnsupported,
                $"OFFICE_CAPABILITY_UNSUPPORTED: {capability}"),
        };
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
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(output))!);
        if (ir.TrimStart().StartsWith('{'))
        {
            // Inline IR JSON: land it in a temp file for the Open XML worker.
            string irTemp = Path.Combine(Path.GetTempPath(), $"deck-ir-{Guid.NewGuid():N}.json");
            File.WriteAllText(irTemp, ir);
            try
            {
                PptxWriter.CompileDeck(irTemp, output);
            }
            finally
            {
                File.Delete(irTemp);
            }
        }
        else
        {
            PptxWriter.CompileDeck(ir, output);
        }
        var info = PptxInspector.Inspect(output);
        return new
        {
            operation_id = Guid.NewGuid().ToString("N"),
            changed = new { files = 1, slides = info.SlideCount },
            warnings = Array.Empty<string>(),
            artefacts = new[] { Artifact(output, "pptx") },
            validation = new { output_openable = true, non_empty = true, slide_count_reasonable = info.SlideCount >= 1 },
            backend = "openxml",
            indeterminate = false,
        };
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
            var outcome = Com().Inspect(path);
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

    private object BatchConvert(JsonElement input)
    {
        var com = Com();
        string targetFormat = input.TryGetProperty("target_format", out var tf) ? tf.GetString() ?? "pdf" : "pdf";
        if (!targetFormat.Equals("pdf", StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeComException(OfficeErrorCode.OfficeCapabilityUnsupported,
                $"target_format '{targetFormat}' is not supported (pdf only)");
        }
        string outputDir = input.TryGetProperty("output_directory", out var od) && od.ValueKind == JsonValueKind.String
            ? od.GetString()!
            : throw new OfficeArgumentException("input.output_directory is required");
        Directory.CreateDirectory(outputDir);

        var paths = ResolveInputs(input);
        if (paths.Length == 0)
        {
            throw new OfficeArgumentException("input.inputs resolved to no files");
        }

        string expected = ExpectedExtension(_app);
        var items = new JsonArray();
        var artefacts = new JsonArray();
        var warnings = new JsonArray();
        int succeeded = 0;
        foreach (var path in paths)
        {
            var ext = Path.GetExtension(path);
            if (!ext.Equals(expected, StringComparison.OrdinalIgnoreCase))
            {
                items.Add(new JsonObject
                {
                    ["input_path"] = path,
                    ["ok"] = false,
                    ["error_code"] = "OFFICE_CAPABILITY_UNSUPPORTED",
                    ["error"] = $"{_app} sidecar handles {expected} only; route '{path}' to its own sidecar",
                });
                continue;
            }
            string outputPath = VersionedOutputPath(outputDir, Path.GetFileNameWithoutExtension(path) + ".pdf");
            var outcome = com.ConvertToPdf(path, outputPath);
            if (outcome.Ok)
            {
                succeeded++;
                artefacts.Add(Artifact(outputPath, "pdf"));
            }
            else if (outcome.ErrorCode is not null)
            {
                warnings.Add($"{path}: {outcome.Error}");
            }
            items.Add(JsonSerializer.SerializeToNode(outcome)!);
        }
        return new
        {
            operation_id = Guid.NewGuid().ToString("N"),
            changed = new
            {
                files = paths.Length,
                succeeded,
                failed = paths.Length - succeeded,
                target_format = targetFormat,
                items,
            },
            warnings,
            artefacts,
            validation = new
            {
                output_openable = succeeded > 0,
                non_empty = succeeded > 0,
                page_count_reasonable = succeeded > 0,
            },
            backend = "desktop_com",
            indeterminate = false,
        };
    }

    private object BatchReplaceText(JsonElement input)
    {
        var com = Com();
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

        string expected = ExpectedExtension(_app);
        var items = new JsonArray();
        int totalMatched = 0;
        int totalReplaced = 0;
        foreach (var path in paths)
        {
            var ext = Path.GetExtension(path);
            if (!ext.Equals(expected, StringComparison.OrdinalIgnoreCase))
            {
                items.Add(new JsonObject
                {
                    ["path"] = path,
                    ["error"] = $"{_app} sidecar handles {expected} only",
                });
                continue;
            }
            var outcome = com.ReplaceText(path, rules, scope, dryRun);
            totalMatched += outcome.TotalMatched;
            totalReplaced += outcome.TotalReplaced;
            items.Add(JsonSerializer.SerializeToNode(outcome)!);
        }
        return new
        {
            operation_id = Guid.NewGuid().ToString("N"),
            changed = new
            {
                files = paths.Length,
                dry_run = dryRun,
                total_matched = totalMatched,
                total_replaced = totalReplaced,
                items,
            },
            warnings = Array.Empty<string>(),
            artefacts = Array.Empty<object>(),
            validation = new { },
            backend = "desktop_com",
            indeterminate = false,
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
            _comAttached = _com.TryAttach(AttachBudget);
        }
    }

    private object Manifest()
    {
        var capabilities = new SortedDictionary<string, string>
        {
            ["deck.compile"] = "0.1.0",
            ["document.inspect"] = "0.1.0",
        };
        var modes = new List<string> { "openxml" };
        object? application = null;
        if (_comAttached)
        {
            modes.Add("desktop_com");
            capabilities["batch.convert"] = "0.1.0";
            capabilities["batch.replace_text"] = "0.1.0";
            if (_app.Equals("powerpoint", StringComparison.OrdinalIgnoreCase))
            {
                capabilities["slide.render"] = "0.1.0";
            }
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
            provider = "dcc-mcp-office",
            provider_version = ProviderVersion,
            protocol_version = ProtocolVersion,
            application,
            execution_modes = modes,
            capabilities,
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

    private static string VersionedOutputPath(string directory, string fileName)
    {
        string candidate = Path.Combine(directory, fileName);
        if (!File.Exists(candidate))
        {
            return candidate;
        }
        string stem = Path.GetFileNameWithoutExtension(fileName);
        string ext = Path.GetExtension(fileName);
        for (int n = 1; ; n++)
        {
            candidate = Path.Combine(directory, $"{stem}-{n}{ext}");
            if (!File.Exists(candidate))
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

    private static string Sha256(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

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
                    paths.Add(spec);
                }
            }
        }
        return paths.Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
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
