using System.Text.Json;
using System.Text.RegularExpressions;

namespace Office.Automation.Host;

/// <summary>
/// Validator for the deliberately small draft-07 assertion subset used by the
/// embedded Office capability schemas. Annotation keywords such as title,
/// description, default, and $id do not affect validation.
/// </summary>
internal static class JsonSchemaValidator
{
    internal static void Validate(JsonElement value, JsonElement schema, string path)
    {
        if (schema.TryGetProperty("type", out JsonElement type))
        {
            ValidateType(value, type, path);
        }
        if (schema.TryGetProperty("enum", out JsonElement allowed)
            && !allowed.EnumerateArray().Any(candidate => candidate.GetRawText() == value.GetRawText()))
        {
            throw Invalid(path, $"must be one of {allowed.GetRawText()}");
        }

        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
                ValidateObject(value, schema, path);
                break;
            case JsonValueKind.Array:
                ValidateArray(value, schema, path);
                break;
            case JsonValueKind.String:
                ValidateString(value.GetString() ?? "", schema, path);
                break;
            case JsonValueKind.Number:
                ValidateNumber(value, schema, path);
                break;
        }
    }

    private static void ValidateType(JsonElement value, JsonElement type, string path)
    {
        string[] accepted = type.ValueKind switch
        {
            JsonValueKind.String => [type.GetString() ?? ""],
            JsonValueKind.Array => type.EnumerateArray()
                .Select(item => item.GetString() ?? "")
                .ToArray(),
            _ => throw new InvalidDataException("JSON Schema type must be a string or array"),
        };
        if (accepted.Length == 0 || !accepted.Any(name => TypeMatches(value, name)))
        {
            throw Invalid(path, $"must be {string.Join(" or ", accepted)}");
        }
    }

    private static bool TypeMatches(JsonElement value, string type) => type switch
    {
        "object" => value.ValueKind == JsonValueKind.Object,
        "array" => value.ValueKind == JsonValueKind.Array,
        "string" => value.ValueKind == JsonValueKind.String,
        "integer" => value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out _),
        "number" => value.ValueKind == JsonValueKind.Number,
        "boolean" => value.ValueKind is JsonValueKind.True or JsonValueKind.False,
        "null" => value.ValueKind == JsonValueKind.Null,
        _ => throw new InvalidDataException($"unsupported JSON Schema type '{type}'"),
    };

    private static void ValidateObject(JsonElement value, JsonElement schema, string path)
    {
        JsonElement properties = schema.TryGetProperty("properties", out JsonElement declared)
            ? declared
            : default;
        if (schema.TryGetProperty("required", out JsonElement required))
        {
            foreach (JsonElement name in required.EnumerateArray())
            {
                string propertyName = name.GetString() ?? "";
                if (!value.TryGetProperty(propertyName, out _))
                {
                    throw Invalid($"{path}.{propertyName}", "is required");
                }
            }
        }
        foreach (JsonProperty property in value.EnumerateObject())
        {
            if (properties.ValueKind == JsonValueKind.Object
                && properties.TryGetProperty(property.Name, out JsonElement propertySchema))
            {
                Validate(property.Value, propertySchema, $"{path}.{property.Name}");
            }
            else if (schema.TryGetProperty("additionalProperties", out JsonElement additional)
                && additional.ValueKind == JsonValueKind.False)
            {
                throw Invalid($"{path}.{property.Name}", "is not allowed");
            }
        }
    }

    private static void ValidateArray(JsonElement value, JsonElement schema, string path)
    {
        int count = value.GetArrayLength();
        if (schema.TryGetProperty("minItems", out JsonElement minimum)
            && count < minimum.GetInt32())
        {
            throw Invalid(path, $"must contain at least {minimum.GetInt32()} item(s)");
        }
        if (schema.TryGetProperty("items", out JsonElement itemSchema))
        {
            int index = 0;
            foreach (JsonElement item in value.EnumerateArray())
            {
                Validate(item, itemSchema, $"{path}[{index}]");
                index++;
            }
        }
    }

    private static void ValidateString(string value, JsonElement schema, string path)
    {
        if (schema.TryGetProperty("minLength", out JsonElement minimum)
            && value.Length < minimum.GetInt32())
        {
            throw Invalid(path, $"must contain at least {minimum.GetInt32()} character(s)");
        }
        if (schema.TryGetProperty("pattern", out JsonElement pattern)
            && !Regex.IsMatch(value, pattern.GetString() ?? ""))
        {
            throw Invalid(path, $"must match {pattern.GetString()}");
        }
    }

    private static void ValidateNumber(JsonElement value, JsonElement schema, string path)
    {
        if (schema.TryGetProperty("minimum", out JsonElement minimum)
            && value.GetDouble() < minimum.GetDouble())
        {
            throw Invalid(path, $"must be at least {minimum.GetRawText()}");
        }
        if (schema.TryGetProperty("maximum", out JsonElement maximum)
            && value.GetDouble() > maximum.GetDouble())
        {
            throw Invalid(path, $"must be at most {maximum.GetRawText()}");
        }
    }

    private static OfficeArgumentException Invalid(string path, string reason) =>
        new($"{path} {reason}");
}
