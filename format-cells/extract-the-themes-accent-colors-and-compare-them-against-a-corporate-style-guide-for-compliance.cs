// Title: Check Excel workbook theme accent colors for corporate style guide compliance using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that reads all six ThemeColorType accent colors from a workbook and prints their hex values. | Enhance the code to compare each extracted accent color with a predefined corporate color array and output a compliance flag for each. | Add robust error handling for missing files or unavailable theme colors and generate a concise compliance report.
// Common Searches: C# Aspose.Cells how to read theme accent colors from an .xlsx file | compare Excel theme palette with corporate brand colors using .NET | validate workbook theme colors against a corporate style guide programmatically | Aspose.Cells GetThemeColor example for compliance checking | generate compliance report for Excel theme colors in C#
// Tags: Aspose.Cells GetThemeColor API | Excel theme palette compliance check | C# corporate color validation for Excel | retrieve workbook theme accents .NET | theme color compliance reporting

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, extracts the six standard theme accent colors via Workbook.GetThemeColor, compares each to a corporate color palette, and prints the hex code with a compliance or non‑compliant status, handling missing files and unavailable theme colors gracefully.
class ThemeComplianceChecker
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the theme accent types (six standard accent colors)
            ThemeColorType[] accentTypes = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            // Define the corporate style guide accent colors
            Color[] corporateColors = new Color[]
            {
                Color.FromArgb(0, 112, 192),   // Corporate Blue
                Color.FromArgb(255, 0, 0),     // Corporate Red
                Color.FromArgb(0, 176, 80),    // Corporate Green
                Color.FromArgb(255, 192, 0),   // Corporate Orange
                Color.FromArgb(112, 48, 160),  // Corporate Purple
                Color.FromArgb(0, 176, 240)    // Corporate Cyan
            };

            // Compare each theme accent color with the corporate palette
            for (int i = 0; i < accentTypes.Length; i++)
            {
                Color accent;
                try
                {
                    // Retrieve the accent color from the workbook's theme
                    accent = workbook.GetThemeColor(accentTypes[i]);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to retrieve theme color {accentTypes[i]}: {ex.Message}");
                    continue;
                }

                bool isCompliant = false;
                foreach (Color corporate in corporateColors)
                {
                    if (accent.ToArgb() == corporate.ToArgb())
                    {
                        isCompliant = true;
                        break;
                    }
                }

                // Output the result for each accent color
                Console.WriteLine(
                    $"Accent {i + 1}: #{accent.R:X2}{accent.G:X2}{accent.B:X2} - {(isCompliant ? "Compliant" : "Non‑compliant")}"
                );
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
