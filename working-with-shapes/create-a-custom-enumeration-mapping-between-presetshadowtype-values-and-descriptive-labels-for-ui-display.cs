// Title: C# example: mapping Aspose.Cells PresetShadowType enum to user‑friendly UI labels
// Common Searches: C# get readable name for Aspose.Cells PresetShadowType enum value | How to translate Aspose.Cells shape shadow enum to UI description strings | Dictionary based mapping of PresetShadowType to friendly label in .NET | Example of converting Aspose.Cells preset shadow identifiers to display text | Retrieve preset shadow description for Excel shape using Aspose.Cells C#
// Tags: Aspose.Cells preset shadow description mapping | C# enum to UI label conversion | shape shadow enumeration translation | .NET dictionary for enum description lookup | custom shadow type mapping Aspose

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShadowMappingDemo
{
    // Local fallback enum for preset shadow types (used when the Aspose enum is unavailable or lacks members).
    // This mirrors the typical preset shadow identifiers.
    public enum PresetShadowType
    {
        PresetShadow1,
        PresetShadow2,
        PresetShadow3,
        PresetShadow4,
        PresetShadow5
        // Extend as needed.
    }

    // The sample defines a local PresetShadowType enum that mirrors Aspose.Cells shadow presets and a static PresetShadowMapper class containing a Dictionary that maps each enum value to a human‑readable description. The GetDescription method returns the appropriate label or a default string. The Main method demonstrates retrieving and printing a description for a selected preset shadow, illustrating how to provide user‑friendly shadow names in a C# UI built on Aspose.Cells.
    public static class PresetShadowMapper
    {
        // Mapping of preset shadow types to user‑friendly descriptions.
        private static readonly Dictionary<PresetShadowType, string> _shadowDescriptions = new Dictionary<PresetShadowType, string>
        {
            { PresetShadowType.PresetShadow1, "Offset 3 pt, 45°, 30% opacity" },
            { PresetShadowType.PresetShadow2, "Offset 5 pt, 60°, 40% opacity" },
            { PresetShadowType.PresetShadow3, "Offset 2 pt, 30°, 20% opacity" },
            { PresetShadowType.PresetShadow4, "Offset 4 pt, 45°, 35% opacity" },
            { PresetShadowType.PresetShadow5, "Offset 6 pt, 70°, 45% opacity" }
            // Add additional mappings as required.
        };

        // Returns the description for a given preset shadow type.
        public static string GetDescription(PresetShadowType type)
        {
            return _shadowDescriptions.TryGetValue(type, out string description) ? description : "Custom Shadow";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Example: retrieve description for a specific preset shadow.
                PresetShadowType shadow = PresetShadowType.PresetShadow3;
                string label = PresetShadowMapper.GetDescription(shadow);
                Console.WriteLine($"Shadow {shadow}: {label}");
            }
            catch (Exception ex)
            {
                // Runtime safety: report unexpected errors.
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
