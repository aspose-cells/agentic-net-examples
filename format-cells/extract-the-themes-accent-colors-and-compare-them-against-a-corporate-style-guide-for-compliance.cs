// Title: C# – Extract Excel Theme Accent Colors and Validate Against Corporate Style Guide with Aspose.Cells
// Description: Loads an Excel workbook, retrieves the six theme accent colors via Aspose.Cells GetThemeColor, compares each ARGB value to a predefined corporate palette, reports compliance per accent, and saves the file unchanged.
// Keywords: Aspose.Cells | C# | GetThemeColor | ThemeColorType | Excel theme colors | corporate color compliance | color palette validation | branding check | programmatic theme extraction
// Common Searches: Aspose.Cells GetThemeColor example C# | How to read Excel theme accent colors in .NET | Validate Excel workbook theme colors against corporate palette | Check Excel theme compliance programmatically | C# code to compare Excel theme colors
// Developer Intent: Identify the workbook’s theme accent colors and determine if they match the organization’s approved color scheme.
// Use Cases: Generate a compliance report for multiple workbooks by iterating through each ThemeColorType and logging mismatches. | Automate brand enforcement by flagging or rejecting Excel files whose theme colors differ from the approved corporate guide. | Integrate the check into a CI/CD pipeline to ensure generated reports use mandated theme colors before release.
// AI Prompts: Write a C# method that returns a list of non‑compliant theme accent colors given a Workbook and an array of corporate colors using Aspose.Cells. | Provide code to batch‑process a folder of Excel files, checking each file’s theme accent colors against a corporate style guide and outputting a summary CSV. | Explain how to add a tolerance range for color differences when validating theme compliance with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeComplianceCheck
{
    // Loads an Excel workbook, retrieves the six theme accent colors via Aspose.Cells GetThemeColor, compares each ARGB value to a predefined corporate palette, reports compliance per accent, and saves the file unchanged.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            Workbook workbook = new Workbook("input.xlsx");

            // Corporate style guide accent colors (example values)
            Color[] corporateAccentColors = new Color[6]
            {
                Color.FromArgb(255, 0, 112, 192),   // Accent1
                Color.FromArgb(255, 255, 192, 0),   // Accent2
                Color.FromArgb(255, 112, 173, 71),  // Accent3
                Color.FromArgb(255, 237, 125, 49),  // Accent4
                Color.FromArgb(255, 191, 0, 0),     // Accent5
                Color.FromArgb(255, 112, 48, 160)   // Accent6
            };

            // Array of ThemeColorType values for accents
            ThemeColorType[] accentTypes = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            // Check each accent color against the corporate guide
            for (int i = 0; i < accentTypes.Length; i++)
            {
                ThemeColorType type = accentTypes[i];
                Color workbookColor = workbook.GetThemeColor(type);
                Color corporateColor = corporateAccentColors[i];

                bool isMatch = workbookColor.ToArgb() == corporateColor.ToArgb();

                Console.WriteLine($"{type}: Workbook Color = {workbookColor}, Corporate Expected = {corporateColor} => {(isMatch ? "Compliant" : "Non‑compliant")}");
            }

            // Save the workbook (no modifications made, just demonstrating lifecycle)
            workbook.Save("output.xlsx");
        }
    }
}
