// Title: C# – Map Aspose.Cells PresetShadowType enum to readable UI labels
// Description: Shows how to build a Dictionary<PresetShadowType,string> that pairs each PresetShadowType value with a friendly description, apply a preset shadow to a rectangle shape in a workbook, fetch the description for UI or logging, and save the workbook as PresetShadowMappingDemo.xlsx.
// Keywords: Aspose.Cells | PresetShadowType | enum to string mapping | C# dictionary | shape shadow description | UI label mapping | Excel workbook example | rectangle shape shadow | Aspose.Cells API demo | code sample
// Common Searches: Aspose.Cells map PresetShadowType to text | C# dictionary for shadow enum descriptions | display friendly shadow names in Excel workbook | retrieve shape shadow label Aspose.Cells | preset shadow type lookup table C#
// Developer Intent: Create a lookup that converts PresetShadowType values into human‑readable strings for UI or logging.
// Use Cases: Populate a drop‑down list with descriptive shadow names for end‑users. | Log the applied shadow effect in a status panel or audit trail. | Synchronize UI controls with a shape's current shadow setting during editing.
// AI Prompts: Generate a method that returns the description for any PresetShadowType using the provided dictionary. | Show how to bind the shadowTypeDescriptions dictionary to a WinForms ComboBox for selecting shadow effects. | Write unit tests that verify each PresetShadowType maps to the correct description string.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace PresetShadowMappingDemo
{
    // Shows how to build a Dictionary<PresetShadowType,string> that pairs each PresetShadowType value with a friendly description, apply a preset shadow to a rectangle shape in a workbook, fetch the description for UI or logging, and save the workbook as PresetShadowMappingDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a mapping between PresetShadowType enum values and user‑friendly descriptions.
            var shadowTypeDescriptions = new Dictionary<PresetShadowType, string>
            {
                { PresetShadowType.NoShadow, "No shadow" },
                { PresetShadowType.Custom, "Custom shadow" },
                { PresetShadowType.OffsetDiagonalBottomRight, "Outer shadow offset diagonal bottom right" },
                { PresetShadowType.OffsetBottom, "Outer shadow offset bottom" },
                { PresetShadowType.OffsetDiagonalBottomLeft, "Outer shadow offset diagonal bottom left" },
                { PresetShadowType.OffsetRight, "Outer shadow offset right" },
                { PresetShadowType.OffsetCenter, "Outer shadow offset center" },
                { PresetShadowType.OffsetLeft, "Outer shadow offset left" },
                { PresetShadowType.OffsetDiagonalTopRight, "Outer shadow offset diagonal top right" },
                { PresetShadowType.OffsetTop, "Outer shadow offset top" },
                { PresetShadowType.OffsetDiagonalTopLeft, "Outer shadow offset diagonal top left" },
                { PresetShadowType.InsideDiagonalTopLeft, "Inner shadow inside diagonal top left" },
                { PresetShadowType.InsideTop, "Inner shadow inside top" },
                { PresetShadowType.InsideDiagonalTopRight, "Inner shadow inside diagonal top right" },
                { PresetShadowType.InsideLeft, "Inner shadow inside left" },
                { PresetShadowType.InsideCenter, "Inner shadow inside center" },
                { PresetShadowType.InsideRight, "Inner shadow inside right" },
                { PresetShadowType.InsideDiagonalBottomLeft, "Inner shadow inside diagonal bottom left" },
                { PresetShadowType.InsideBottom, "Inner shadow inside bottom" },
                { PresetShadowType.InsideDiagonalBottomRight, "Inner shadow inside diagonal bottom right" },
                { PresetShadowType.PerspectiveDiagonalUpperLeft, "Outer shadow perspective diagonal upper left" },
                { PresetShadowType.PerspectiveDiagonalUpperRight, "Outer shadow perspective diagonal upper right" },
                { PresetShadowType.Below, "Outer shadow below" },
                { PresetShadowType.PerspectiveDiagonalLowerLeft, "Outer shadow perspective diagonal lower left" },
                { PresetShadowType.PerspectiveDiagonalLowerRight, "Outer shadow perspective diagonal lower right" }
            };

            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate a shadow effect
            Shape rect = sheet.Shapes.AddRectangle(2, 2, 100, 100, 200, 200);

            // Choose a preset shadow type
            rect.ShadowEffect.PresetType = PresetShadowType.OffsetDiagonalBottomRight;

            // Retrieve the description for UI display
            string description = shadowTypeDescriptions[rect.ShadowEffect.PresetType];
            Console.WriteLine($"Applied shadow type: {rect.ShadowEffect.PresetType} – {description}");

            // Save the workbook (lifecycle: save)
            workbook.Save("PresetShadowMappingDemo.xlsx");
        }
    }
}
