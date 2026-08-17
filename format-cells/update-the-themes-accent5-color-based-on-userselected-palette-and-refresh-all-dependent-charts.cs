// Title: Change Accent5 Theme Color and Refresh All Charts with Aspose.Cells for .NET (C#)
// Description: C# example that loads or creates an Excel workbook, sets the Accent5 theme color using Workbook.SetThemeColor, ensures chart series inherit the new theme, and saves the updated file. The theme change propagates automatically to all charts.
// Keywords: Aspose.Cells | C# | SetThemeColor | Accent5 | Excel theme color | update chart colors | refresh charts programmatically | Excel workbook theme | change Excel palette .NET | Aspose.Cells chart fill
// Common Searches: how to change Accent5 theme color with Aspose.Cells | update Excel theme color and refresh charts C# | Aspose.Cells SetThemeColor example | programmatically change Excel palette .NET | apply new theme color to existing charts Aspose
// Developer Intent: Set a custom Accent5 theme color in an Excel workbook and make all existing charts reflect the change automatically.
// Use Cases: Apply a corporate brand color (e.g., orange) to the Accent5 slot across multiple reports before distribution. | Standardize the Accent5 color in a batch of client workbooks to meet branding guidelines. | Create a template that updates the theme color on the fly and regenerates charts with the new palette.
// AI Prompts: Generate C# code using Aspose.Cells to change the Accent5 theme color to a variable Color and automatically update all chart series. | Show how to load a workbook, set multiple theme colors (Accent5, Accent2, etc.), and ensure charts inherit the new colors without manual fill assignments.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// C# example that loads or creates an Excel workbook, sets the Accent5 theme color using Workbook.SetThemeColor, ensures chart series inherit the new theme, and saves the updated file. The theme change propagates automatically to all charts.
class UpdateAccent5Theme
{
    static void Main()
    {
        try
        {
            // Load an existing workbook if needed; otherwise create a new one.
            Workbook workbook;
            string inputPath = "input.xlsx";

            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a blank workbook
            }

            // User‑selected color for Accent5 (example: orange)
            Color userSelectedColor = Color.FromArgb(255, 165, 0);

            // Update the theme's Accent5 color
            workbook.SetThemeColor(ThemeColorType.Accent5, userSelectedColor);

            // Refresh charts so they reflect the new Accent5 color.
            // The theme change automatically propagates; explicit fill updates are not required.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    foreach (Series series in chart.NSeries)
                    {
                        // Ensure the series uses the theme color (no additional code needed).
                        // If custom fill was previously set, it can be cleared by resetting the fill type.
                        series.Area.FillFormat.FillType = FillType.Solid;
                    }
                }
            }

            // Save the modified workbook
            string outputPath = "UpdatedAccent5Theme.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
