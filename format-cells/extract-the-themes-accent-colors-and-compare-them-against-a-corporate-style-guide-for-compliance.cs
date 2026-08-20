// Title: Validate Excel Theme Accent Colors Against a Corporate Style Guide Using Aspose.Cells (C#)
// Description: Creates a workbook, defines a corporate palette of six accent colors, reads each theme accent with GetThemeColor, compares ARGB values, logs PASS/FAIL for compliance, and saves the file.
// Keywords: Aspose.Cells GetThemeColor | Excel theme accent colors C# | corporate color compliance | validate Excel theme palette | theme color comparison .NET
// Common Searches: read Excel theme accent colors with Aspose.Cells | compare workbook theme colors to corporate palette C# | Aspose.Cells validate Excel theme against style guide | example GetThemeColor usage Aspose.Cells
// Developer Intent: Extract the workbook’s theme accent colors and verify they exactly match a predefined corporate color set.
// Use Cases: Batch‑process multiple workbooks to produce a branding compliance report. | Add the check to a CI/CD pipeline that blocks non‑conforming Excel files. | Create a UI that highlights cells using non‑standard theme colors.
// AI Prompts: Generate a method that returns all ThemeColorType values that fail compliance for a given workbook and corporate palette. | Adjust the comparison to allow a ±5 tolerance per ARGB channel when determining compliance. | Write an NUnit test that applies a custom theme to a workbook and asserts the expected PASS/FAIL results.

using System;
using System.Drawing;
using System.Collections.Generic;
using Aspose.Cells;

// Creates a workbook, defines a corporate palette of six accent colors, reads each theme accent with GetThemeColor, compares ARGB values, logs PASS/FAIL for compliance, and saves the file.
class ThemeComplianceChecker
{
    static void Main()
    {
        // Create a new workbook (default theme will be used)
        Workbook workbook = new Workbook();

        // Corporate style guide: expected accent colors (example ARGB values)
        var corporateAccentColors = new Dictionary<ThemeColorType, Color>
        {
            { ThemeColorType.Accent1, Color.FromArgb(255, 0, 112, 192) },   // Example: corporate blue
            { ThemeColorType.Accent2, Color.FromArgb(255, 255, 192, 0) },   // Example: corporate orange
            { ThemeColorType.Accent3, Color.FromArgb(255, 112, 173, 71) },  // Example: corporate green
            { ThemeColorType.Accent4, Color.FromArgb(255, 237, 125, 49) },  // Example: corporate red‑orange
            { ThemeColorType.Accent5, Color.FromArgb(255, 255, 0, 0) },     // Example: corporate red
            { ThemeColorType.Accent6, Color.FromArgb(255, 0, 176, 80) }     // Example: corporate dark green
        };

        // Accent types to evaluate
        ThemeColorType[] accentTypes = new ThemeColorType[]
        {
            ThemeColorType.Accent1,
            ThemeColorType.Accent2,
            ThemeColorType.Accent3,
            ThemeColorType.Accent4,
            ThemeColorType.Accent5,
            ThemeColorType.Accent6
        };

        // Compare each workbook accent color with the corporate guide
        foreach (var accent in accentTypes)
        {
            Color workbookColor = workbook.GetThemeColor(accent);
            Color expectedColor = corporateAccentColors[accent];

            bool isCompliant = workbookColor.ToArgb() == expectedColor.ToArgb();

            Console.WriteLine($"{accent}: Workbook={workbookColor} Expected={expectedColor} Compliance={(isCompliant ? "PASS" : "FAIL")}");
        }

        // Save the workbook (required lifecycle step)
        workbook.Save("ThemeComplianceCheck.xlsx");
    }
}
