using System.Text.RegularExpressions;

namespace Office.Automation.Host;

public sealed record InputResolution(
    IReadOnlyList<string> Paths,
    IReadOnlyList<string> Warnings);

/// <summary>Resolves literals and bounded globs without implicit recursion.</summary>
public static class InputResolver
{
    public static InputResolution Resolve(IEnumerable<string> specifications)
    {
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var warnings = new List<string>();
        foreach (string specification in specifications)
        {
            try
            {
                ResolveSpecification(specification, paths, warnings);
            }
            catch (Exception error) when (error is ArgumentException
                or NotSupportedException
                or PathTooLongException)
            {
                warnings.Add($"invalid input specification: {specification}");
            }
        }
        return new InputResolution(
            paths.Order(StringComparer.OrdinalIgnoreCase).ToArray(),
            warnings);
    }

    private static void ResolveSpecification(
        string specification,
        ISet<string> paths,
        ICollection<string> warnings)
    {
        if (string.IsNullOrWhiteSpace(specification))
        {
            warnings.Add("input specification was empty");
            return;
        }
        if (!specification.Contains('*') && !specification.Contains('?'))
        {
            string literal = Path.GetFullPath(specification);
            if (File.Exists(literal))
            {
                paths.Add(literal);
            }
            else
            {
                warnings.Add($"input did not match a file: {specification}");
            }
            return;
        }

        string[] matches = ExpandGlob(specification, warnings).ToArray();
        if (matches.Length == 0)
        {
            warnings.Add($"input glob matched no files: {specification}");
        }
        foreach (string match in matches)
        {
            paths.Add(Path.GetFullPath(match));
        }
    }

    private static IEnumerable<string> ExpandGlob(
        string pattern,
        ICollection<string> warnings)
    {
        string normalized = Path.GetFullPath(pattern).Replace('/', '\\');
        int wildcardAt = normalized.IndexOfAny(['*', '?']);
        string literalPrefix = normalized[..wildcardAt];
        string root = Directory.Exists(literalPrefix)
            ? literalPrefix
            : Path.GetDirectoryName(literalPrefix) ?? Path.GetPathRoot(normalized)!;
        if (!Directory.Exists(root))
        {
            return Array.Empty<string>();
        }

        var regex = new Regex(
            "^" + Regex.Escape(normalized)
                .Replace(@"\*\*\\", @"(?:.*\\)?")
                .Replace(@"\*\*", ".*")
                .Replace(@"\*", @"[^\\]*")
                .Replace(@"\?", @"[^\\]") + "$",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        bool recursive = normalized.Contains("**", StringComparison.Ordinal);
        return EnumerateFiles(root, recursive, warnings)
            .Where(path => regex.IsMatch(path));
    }

    private static IEnumerable<string> EnumerateFiles(
        string root,
        bool recursive,
        ICollection<string> warnings)
    {
        var pending = new Queue<string>();
        pending.Enqueue(root);
        while (pending.Count > 0)
        {
            string directory = pending.Dequeue();
            string[] files;
            try
            {
                files = Directory.GetFiles(directory);
            }
            catch (Exception error) when (error is UnauthorizedAccessException or IOException)
            {
                warnings.Add($"skipped inaccessible directory: {directory}");
                continue;
            }
            foreach (string file in files)
            {
                yield return file;
            }
            if (!recursive)
            {
                continue;
            }

            string[] children;
            try
            {
                children = Directory.GetDirectories(directory);
            }
            catch (Exception error) when (error is UnauthorizedAccessException or IOException)
            {
                warnings.Add($"skipped inaccessible directory: {directory}");
                continue;
            }
            foreach (string child in children)
            {
                try
                {
                    if ((File.GetAttributes(child) & FileAttributes.ReparsePoint) == 0)
                    {
                        pending.Enqueue(child);
                    }
                }
                catch (Exception error) when (error is UnauthorizedAccessException or IOException)
                {
                    warnings.Add($"skipped inaccessible directory: {child}");
                }
            }
        }
    }
}
