// Title: Map Aspose.Cells PresetShadowType enum to user‑friendly labels in C#
// Common Searches: how to get a readable name for a PresetShadowType value in Aspose.Cells C# | C# dictionary lookup for Aspose.Cells preset shadow enum descriptions | convert Aspose.Cells PresetShadowType to display text for UI | example of mapping Aspose.Cells shadow types to friendly strings in .NET
// Tags: Aspose.Cells preset shadow enum to label mapping | C# enum description dictionary for Aspose.Cells | friendly shadow type names for UI | custom shadow label lookup in .NET | display Aspose.Cells shadow names in application

using System;
using System.Collections.Generic;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Local definition of preset shadow types matching Aspose.Cells drawing capabilities.
    // If a newer Aspose.Cells version provides its own PresetShadowType enum, replace this
    // definition with the library enum and remove the duplicate.
    public enum PresetShadowType
    {
        PresetShadow1, PresetShadow2, PresetShadow3, PresetShadow4,
        PresetShadow5, PresetShadow6, PresetShadow7, PresetShadow8,
        PresetShadow9, PresetShadow10, PresetShadow11, PresetShadow12,
        PresetShadow13, PresetShadow14, PresetShadow15, PresetShadow16,
        PresetShadow17, PresetShadow18, PresetShadow19, PresetShadow20,
        PresetShadow21, PresetShadow22, PresetShadow23, PresetShadow24,
        PresetShadow25, PresetShadow26, PresetShadow27, PresetShadow28,
        PresetShadow29, PresetShadow30, PresetShadow31, PresetShadow32,
        PresetShadow33, PresetShadow34, PresetShadow35, PresetShadow36
    }

    // The sample defines a local PresetShadowType enumeration that mirrors Aspose.Cells shadow options and provides a static ShadowTypeMapper class containing a Dictionary that maps each enum value to a human‑readable description. The GetLabel method returns the friendly label or falls back to the enum name, and the example program demonstrates retrieving a label for a specific shadow type.
    public static class ShadowTypeMapper
    {
        // Mapping between preset shadow enum values and user‑friendly descriptions.
        private static readonly Dictionary<PresetShadowType, string> _shadowLabels = new Dictionary<PresetShadowType, string>
        {
            { PresetShadowType.PresetShadow1, "Outer Shadow – Soft Edge" },
            { PresetShadowType.PresetShadow2, "Outer Shadow – Hard Edge" },
            { PresetShadowType.PresetShadow3, "Inner Shadow – Soft Edge" },
            { PresetShadowType.PresetShadow4, "Inner Shadow – Hard Edge" },
            { PresetShadowType.PresetShadow5, "Perspective Shadow – Light" },
            { PresetShadowType.PresetShadow6, "Perspective Shadow – Dark" },
            { PresetShadowType.PresetShadow7, "Diagonal Shadow – Light" },
            { PresetShadowType.PresetShadow8, "Diagonal Shadow – Dark" },
            { PresetShadowType.PresetShadow9, "Top Left Shadow – Light" },
            { PresetShadowType.PresetShadow10, "Top Left Shadow – Dark" },
            { PresetShadowType.PresetShadow11, "Bottom Right Shadow – Light" },
            { PresetShadowType.PresetShadow12, "Bottom Right Shadow – Dark" },
            { PresetShadowType.PresetShadow13, "Top Right Shadow – Light" },
            { PresetShadowType.PresetShadow14, "Top Right Shadow – Dark" },
            { PresetShadowType.PresetShadow15, "Bottom Left Shadow – Light" },
            { PresetShadowType.PresetShadow16, "Bottom Left Shadow – Dark" },
            { PresetShadowType.PresetShadow17, "Center Shadow – Light" },
            { PresetShadowType.PresetShadow18, "Center Shadow – Dark" },
            { PresetShadowType.PresetShadow19, "Outer Shadow – Wide" },
            { PresetShadowType.PresetShadow20, "Outer Shadow – Narrow" },
            { PresetShadowType.PresetShadow21, "Inner Shadow – Wide" },
            { PresetShadowType.PresetShadow22, "Inner Shadow – Narrow" },
            { PresetShadowType.PresetShadow23, "Perspective Shadow – Wide" },
            { PresetShadowType.PresetShadow24, "Perspective Shadow – Narrow" },
            { PresetShadowType.PresetShadow25, "Diagonal Shadow – Wide" },
            { PresetShadowType.PresetShadow26, "Diagonal Shadow – Narrow" },
            { PresetShadowType.PresetShadow27, "Top Left Shadow – Wide" },
            { PresetShadowType.PresetShadow28, "Top Left Shadow – Narrow" },
            { PresetShadowType.PresetShadow29, "Bottom Right Shadow – Wide" },
            { PresetShadowType.PresetShadow30, "Bottom Right Shadow – Narrow" },
            { PresetShadowType.PresetShadow31, "Top Right Shadow – Wide" },
            { PresetShadowType.PresetShadow32, "Top Right Shadow – Narrow" },
            { PresetShadowType.PresetShadow33, "Bottom Left Shadow – Wide" },
            { PresetShadowType.PresetShadow34, "Bottom Left Shadow – Narrow" },
            { PresetShadowType.PresetShadow35, "Center Shadow – Wide" },
            { PresetShadowType.PresetShadow36, "Center Shadow – Narrow" }
        };

        /// <param name="type">The preset shadow type.</param>
        /// <returns>User‑friendly description; if not found, returns the enum name.</returns>
        public static string GetLabel(PresetShadowType type)
        {
            return _shadowLabels.TryGetValue(type, out string label) ? label : type.ToString();
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Example: retrieve label for a specific preset shadow type.
                PresetShadowType shadow = PresetShadowType.PresetShadow5;
                string description = ShadowTypeMapper.GetLabel(shadow);
                Console.WriteLine($"Shadow type {shadow} is displayed as: {description}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
