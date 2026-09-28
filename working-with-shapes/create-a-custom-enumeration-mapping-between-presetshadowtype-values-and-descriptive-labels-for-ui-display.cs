// Title: Create a lazy‑initialized C# read‑only dictionary that maps Aspose.Cells PresetShadowType enum values to user‑friendly shadow labels
// AI Prompts: Generate a static C# class that lazily builds a read‑only dictionary linking every PresetShadowType enum member to a concise UI label. | Write a method that receives a PresetShadowType value and returns its friendly label, providing a default string for undefined entries. | Demonstrate iterating over the mapping to populate a UI control such as a dropdown list with the shadow names.
// Common Searches: how to convert Aspose.Cells PresetShadowType enum to readable text in C# | lazy initialization pattern for enum‑to‑string dictionary in Aspose.Cells | populate a combo box with Aspose.Cells shadow preset names using C# | C# example for mapping PresetShadowType to display strings for UI
// Tags: Aspose.Cells PresetShadowType label mapping | lazy initialized enum‑string dictionary C# | friendly shadow name generation Aspose.Cells | UI dropdown population from PresetShadowType | read‑only mapping of shadow presets

using System;
using System.Collections.Generic;
using Aspose.Cells.Drawing;

// The example defines a static PresetShadowMapper class that lazily creates a read‑only dictionary mapping each PresetShadowType enum value to a user‑friendly label such as "Shadow 1". It provides a Mapping property, a GetLabel method with a fallback for unknown values, and a label generator that strips the "PresetShadow" prefix. A sample program enumerates the enum, retrieves a label, and prints the full mapping, illustrating how to fill a UI dropdown with these labels.
public static class PresetShadowMapper
{
    // Lazy‑initialized dictionary that maps each PresetShadowType to a friendly UI label.
    private static readonly Lazy<IReadOnlyDictionary<PresetShadowType, string>> _lazyMap =
        new Lazy<IReadOnlyDictionary<PresetShadowType, string>>(CreateMapping);

    // Public accessor for the mapping.
    public static IReadOnlyDictionary<PresetShadowType, string> Mapping => _lazyMap.Value;

    // Retrieves the friendly label for a specific PresetShadowType.
    public static string GetLabel(PresetShadowType type)
    {
        return Mapping.TryGetValue(type, out var label) ? label : "Unknown Shadow Type";
    }

    // Builds the dictionary by iterating over all enum values.
    private static IReadOnlyDictionary<PresetShadowType, string> CreateMapping()
    {
        var dict = new Dictionary<PresetShadowType, string>();
        foreach (PresetShadowType type in Enum.GetValues(typeof(PresetShadowType)))
        {
            dict[type] = GenerateLabel(type);
        }
        return dict;
    }

    // Generates a descriptive label from the enum name.
    private static string GenerateLabel(PresetShadowType type)
    {
        // Enum names are like "PresetShadow1", "PresetShadow2", etc.
        // Convert them to a more readable form, e.g., "Shadow 1".
        const string prefix = "PresetShadow";
        var name = type.ToString();
        if (name.StartsWith(prefix, StringComparison.Ordinal))
        {
            var numberPart = name.Substring(prefix.Length);
            return $"Shadow {numberPart}";
        }
        // Fallback for any unexpected naming.
        return name;
    }
}

// Example usage.
class Program
{
    static void Main()
    {
        try
        {
            // Safely obtain the first enum value for demonstration.
            var enumValues = Enum.GetValues(typeof(PresetShadowType));
            if (enumValues.Length == 0)
            {
                Console.WriteLine("No PresetShadowType values are defined.");
                return;
            }

            var firstShadow = (PresetShadowType)enumValues.GetValue(0);
            var label = PresetShadowMapper.GetLabel(firstShadow);
            Console.WriteLine($"Label for {firstShadow}: {label}");

            // Simulate populating a UI dropdown (console output).
            foreach (var kvp in PresetShadowMapper.Mapping)
            {
                Console.WriteLine($"{kvp.Key} => {kvp.Value}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
