using System.Text.Json;
using System.Text.Json.Nodes;
using Office.Automation.Com;

namespace Office.Automation.Host;

/// <summary>
/// Effective second-layer command policy derived from the canonical catalog.
/// A request may tighten a default, but cannot relax it.
/// </summary>
internal sealed class CommandPolicy
{
    private readonly IReadOnlyDictionary<string, string> _actions;

    private CommandPolicy(
        IReadOnlyDictionary<string, string> actions,
        string workspaceRoot,
        JsonObject effectivePolicy)
    {
        _actions = actions;
        WorkspaceRoot = workspaceRoot;
        EffectivePolicy = effectivePolicy;
    }

    internal string WorkspaceRoot { get; }

    internal JsonObject EffectivePolicy { get; }

    internal JsonNode? ConfirmationForAudit { get; private set; }

    internal static CommandPolicy Evaluate(
        JsonElement parameters,
        CatalogSecurityPolicy canonical,
        string hostWorkspaceRoot)
    {
        JsonElement requested = parameters.TryGetProperty("policy", out JsonElement policy)
            ? policy
            : JsonSerializer.SerializeToElement(new { });
        if (requested.ValueKind != JsonValueKind.Object)
        {
            throw new OfficeArgumentException("params.policy must be an object");
        }

        var actions = new Dictionary<string, string>(StringComparer.Ordinal);
        var effective = new JsonObject();
        foreach ((string name, string defaultAction) in canonical.Actions)
        {
            string action = defaultAction;
            if (requested.TryGetProperty(name, out JsonElement supplied))
            {
                if (supplied.ValueKind != JsonValueKind.String)
                {
                    throw new OfficeArgumentException($"policy.{name} must be a string");
                }
                action = supplied.GetString() ?? "";
                bool unchanged = action == defaultAction;
                bool tightenedToDeny = action == "deny" && defaultAction != "deny";
                if (!unchanged && !tightenedToDeny)
                {
                    throw new OfficeComException(
                        PolicyError(name),
                        $"policy.{name} cannot relax canonical action '{defaultAction}' to '{action}'");
                }
            }
            actions.Add(name, action);
            effective[name] = action;
        }

        ValidateKnownProperties(requested, canonical.Actions.Keys);
        ValidateWorkspaceOnly(requested);
        ValidateExecuteMso(requested, canonical);
        bool checkpoint = ReadBoolean(requested, "checkpoint", defaultValue: true);
        if (!checkpoint)
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeCapabilityUnsupported,
                "policy.checkpoint cannot disable the canonical write pre-image requirement");
        }
        bool renderAfter = ReadBoolean(requested, "render_after", defaultValue: false);
        string workspaceRoot = WorkspaceGuard.CanonicalizeRoot(hostWorkspaceRoot);
        if (requested.TryGetProperty("workspace_root", out JsonElement root))
        {
            if (root.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(root.GetString()))
            {
                throw new OfficeArgumentException("policy.workspace_root must be a non-empty path");
            }
            string requestedRoot = WorkspaceGuard.CanonicalizeRoot(root.GetString()!);
            if (!requestedRoot.Equals(workspaceRoot, StringComparison.OrdinalIgnoreCase))
            {
                throw new OfficeComException(
                    OfficeErrorCode.OfficeAccessDenied,
                    "policy.workspace_root cannot replace the workspace bound when the host started");
            }
        }

        effective["workspace_only"] = true;
        effective["workspace_root"] = workspaceRoot;
        effective["execute_mso_allowlist"] = new JsonObject();
        effective["execute_mso_confirm"] = new JsonArray(
            canonical.ExecuteMsoConfirm
                .Select(value => (JsonNode?)JsonValue.Create(value))
                .ToArray());
        effective["checkpoint"] = true;
        effective["render_after"] = renderAfter;
        return new CommandPolicy(actions, workspaceRoot, effective);
    }

    internal void ValidateWorkspace(CapabilityDefinition capability, JsonElement input) =>
        WorkspaceGuard.Validate(capability.HandlerId, input, WorkspaceRoot);

    internal void RequireConfirmation(JsonElement parameters, string action)
    {
        string policyAction = _actions.GetValueOrDefault(action)
            ?? throw new InvalidOperationException($"canonical policy has no '{action}' action");
        if (policyAction == "deny")
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeAccessDenied,
                $"policy.{action} denies this operation");
        }
        if (policyAction is not ("confirm" or "checkpoint_and_confirm" or "deny_or_confirm"))
        {
            return;
        }
        if (!parameters.TryGetProperty("confirmation", out JsonElement confirmation)
            || confirmation.ValueKind != JsonValueKind.Object
            || !confirmation.TryGetProperty("confirmed", out JsonElement confirmed)
            || confirmed.ValueKind != JsonValueKind.True
            || !StringPropertyEquals(confirmation, "action", action)
            || !HasHumanIdentity(confirmation, "confirmed_by")
            || !HasTimestamp(confirmation, "confirmed_at"))
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeUserConfirmationRequired,
                $"action '{action}' requires confirmation with action, confirmed=true, confirmed_by='human:<id>', and confirmed_at");
        }
        ConfirmationForAudit = JsonSerializer.SerializeToNode(confirmation);
    }

    private static void ValidateKnownProperties(
        JsonElement requested,
        IEnumerable<string> actionNames)
    {
        var known = actionNames.ToHashSet(StringComparer.Ordinal);
        known.UnionWith([
            "workspace_only",
            "workspace_root",
            "execute_mso_allowlist",
            "execute_mso_confirm",
            "checkpoint",
            "render_after",
        ]);
        foreach (JsonProperty property in requested.EnumerateObject())
        {
            if (!known.Contains(property.Name))
            {
                throw new OfficeArgumentException($"policy.{property.Name} is not defined by the catalog");
            }
        }
    }

    private static void ValidateWorkspaceOnly(JsonElement requested)
    {
        if (!requested.TryGetProperty("workspace_only", out JsonElement value))
        {
            return;
        }
        if (value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
        {
            throw new OfficeArgumentException("policy.workspace_only must be a boolean");
        }
        if (value.ValueKind == JsonValueKind.False)
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeAccessDenied,
                "policy.workspace_only cannot be disabled at the host boundary");
        }
    }

    private static void ValidateExecuteMso(
        JsonElement requested,
        CatalogSecurityPolicy canonical)
    {
        if (requested.TryGetProperty("execute_mso_allowlist", out JsonElement allowlist))
        {
            if (allowlist.ValueKind != JsonValueKind.Object)
            {
                throw new OfficeArgumentException("policy.execute_mso_allowlist must be an object");
            }
            if (allowlist.EnumerateObject().Any())
            {
                throw new OfficeComException(
                    OfficeErrorCode.OfficeCapabilityUnsupported,
                    "policy.execute_mso_allowlist must stay empty in this host");
            }
        }
        if (requested.TryGetProperty("execute_mso_confirm", out JsonElement confirm))
        {
            if (confirm.ValueKind != JsonValueKind.Array
                || !confirm.EnumerateArray().Select(value => value.GetString())
                    .SequenceEqual(canonical.ExecuteMsoConfirm))
            {
                throw new OfficeComException(
                    OfficeErrorCode.OfficeCapabilityUnsupported,
                    "policy.execute_mso_confirm cannot diverge from the canonical catalog");
            }
        }
    }

    private static bool ReadBoolean(JsonElement policy, string name, bool defaultValue)
    {
        if (!policy.TryGetProperty(name, out JsonElement value))
        {
            return defaultValue;
        }
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            _ => throw new OfficeArgumentException($"policy.{name} must be a boolean"),
        };
    }

    private static bool StringPropertyEquals(JsonElement value, string name, string expected) =>
        value.TryGetProperty(name, out JsonElement property)
        && property.ValueKind == JsonValueKind.String
        && property.GetString() == expected;

    private static bool HasHumanIdentity(JsonElement value, string name) =>
        value.TryGetProperty(name, out JsonElement property)
        && property.ValueKind == JsonValueKind.String
        && property.GetString() is string identity
        && identity.StartsWith("human:", StringComparison.Ordinal)
        && identity.Length > "human:".Length;

    private static bool HasTimestamp(JsonElement value, string name) =>
        value.TryGetProperty(name, out JsonElement property)
        && property.ValueKind == JsonValueKind.String
        && DateTimeOffset.TryParse(property.GetString(), out _);

    private static OfficeErrorCode PolicyError(string name) => name switch
    {
        "vba_application_run" or "macros" or "ole_activex_activation" or "access_macros" =>
            OfficeErrorCode.OfficeMacroBlocked,
        "external_links_auto_update" => OfficeErrorCode.OfficeExternalLinkBlocked,
        "protected_view_bypass" => OfficeErrorCode.OfficeProtectedView,
        _ => OfficeErrorCode.OfficeCapabilityUnsupported,
    };
}

internal static class WorkspaceGuard
{
    internal static string CanonicalizeRoot(string path)
    {
        string root = Canonicalize(path);
        if (!Directory.Exists(root))
        {
            throw new OfficeArgumentException($"workspace root does not exist: {root}");
        }
        return TrimTrailingSeparators(root);
    }

    internal static void Validate(CapabilityHandler handler, JsonElement input, string workspaceRoot)
    {
        IEnumerable<string> paths = handler switch
        {
            CapabilityHandler.DeckCompile => DeckPaths(input),
            CapabilityHandler.DocumentInspect => StringProperty(input, "path"),
            CapabilityHandler.BatchConvert => InputPaths(input)
                .Concat(StringProperty(input, "output_directory")),
            CapabilityHandler.BatchReplaceText => InputPaths(input),
            CapabilityHandler.SlideRender => StringProperty(input, "path")
                .Concat(StringProperty(input, "output_directory")),
            _ => throw new ArgumentOutOfRangeException(nameof(handler)),
        };
        ValidatePaths(paths, workspaceRoot);
    }

    internal static void ValidatePaths(IEnumerable<string> paths, string workspaceRoot)
    {
        foreach (string path in paths)
        {
            EnsureInside(path, workspaceRoot);
        }
    }

    private static IEnumerable<string> DeckPaths(JsonElement input)
    {
        foreach (string output in StringProperty(input, "output"))
        {
            yield return output;
        }
        foreach (string ir in StringProperty(input, "ir"))
        {
            if (!ir.TrimStart().StartsWith('{'))
            {
                yield return ir;
            }
        }
    }

    private static IEnumerable<string> InputPaths(JsonElement input)
    {
        if (!input.TryGetProperty("inputs", out JsonElement inputs)
            || inputs.ValueKind != JsonValueKind.Array)
        {
            yield break;
        }
        foreach (JsonElement item in inputs.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.String && item.GetString() is string path)
            {
                yield return path;
            }
        }
    }

    private static IEnumerable<string> StringProperty(JsonElement input, string name)
    {
        if (input.TryGetProperty(name, out JsonElement value)
            && value.ValueKind == JsonValueKind.String
            && value.GetString() is string path)
        {
            yield return path;
        }
    }

    private static void EnsureInside(string path, string workspaceRoot)
    {
        string candidate = path.IndexOfAny(['*', '?']) is int wildcard && wildcard >= 0
            ? Canonicalize(GlobBase(path, wildcard))
            : Canonicalize(path);
        string root = CanonicalizeRoot(workspaceRoot);
        string rootPrefix = root.EndsWith(Path.DirectorySeparatorChar)
            || root.EndsWith(Path.AltDirectorySeparatorChar)
            ? root
            : root + Path.DirectorySeparatorChar;
        if (!candidate.Equals(root, StringComparison.OrdinalIgnoreCase)
            && !candidate.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeAccessDenied,
                $"path '{candidate}' is outside workspace '{root}'");
        }
    }

    private static string GlobBase(string pattern, int wildcard)
    {
        string prefix = pattern[..wildcard];
        string directory = prefix.EndsWith(Path.DirectorySeparatorChar)
            || prefix.EndsWith(Path.AltDirectorySeparatorChar)
            ? prefix
            : Path.GetDirectoryName(prefix) ?? Directory.GetCurrentDirectory();
        return Path.GetFullPath(directory);
    }

    /// <summary>
    /// Resolves every existing reparse-point component so a junction or
    /// symbolic link inside the workspace cannot redirect an operation out
    /// of the process-bound root. Non-existent output suffixes stay lexical.
    /// </summary>
    private static string Canonicalize(string path)
    {
        try
        {
            string full = Path.GetFullPath(path);
            string volumeRoot = Path.GetPathRoot(full)
                ?? throw new OfficeArgumentException($"path has no volume root: {path}");
            string current = volumeRoot;
            string remainder = full[volumeRoot.Length..];
            foreach (string component in remainder.Split(
                [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
                StringSplitOptions.RemoveEmptyEntries))
            {
                string next = Path.Combine(current, component);
                FileSystemInfo? entry = Directory.Exists(next)
                    ? new DirectoryInfo(next)
                    : File.Exists(next)
                        ? new FileInfo(next)
                        : null;
                FileSystemInfo? target = entry?.ResolveLinkTarget(returnFinalTarget: true);
                current = Path.GetFullPath(target?.FullName ?? next);
            }
            return TrimTrailingSeparators(current);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new OfficeComException(
                OfficeErrorCode.OfficeAccessDenied,
                $"path could not be resolved safely: {path}: {ex.Message}",
                ex);
        }
    }

    private static string TrimTrailingSeparators(string path)
    {
        string volumeRoot = Path.GetPathRoot(path) ?? "";
        return path.Length <= volumeRoot.Length
            ? path
            : path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}
