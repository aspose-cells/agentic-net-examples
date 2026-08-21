// Title: C# Aspose.Cells: Change Accent6 Theme to Pastel LightPink and Report Chart Legend Positions
// Description: Loads an existing workbook, applies a pastel LightPink shade to the Accent6 theme color with SetThemeColor, iterates through every worksheet and chart to output each legend's position, and saves the updated file.
// Keywords: Aspose.Cells SetThemeColor Accent6 | C# change Excel theme color pastel | Aspose.Cells chart legend position | modify workbook theme Aspose.Cells .NET | iterate charts workbook Aspose.Cells
// Common Searches: how to set Accent6 theme color using Aspose.Cells C# | list chart legend positions in all sheets with Aspose.Cells | Aspose.Cells example change theme to pastel color | C# code to update Excel theme and check chart legends
// Developer Intent: Apply a custom pastel color to the workbook’s Accent6 theme and retrieve the position of each chart legend.
// Use Cases: Standardize report colors by updating the Accent6 theme across existing Excel files. | Audit chart legends in multi‑sheet workbooks to ensure consistent placement before publishing. | Create a version of a workbook with a new pastel palette for branding or visual accessibility.
// AI Prompts: Generate C# Aspose.Cells code that sets Accent6 to a specific RGB value and prints every chart legend’s position. | Show how to loop through all worksheets and charts in a workbook with Aspose.Cells and log legend details. | Explain workarounds for checking legend visibility in Aspose.Cells when the IsVisible property is unavailable.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsThemeAndChartDemo
{
    // Loads an existing workbook, applies a pastel LightPink shade to the Accent6 theme color with SetThemeColor, iterates through every worksheet and chart to output each legend's position, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Define a pastel shade for Accent6 (light pink)
                Color pastelAccent6 = Color.FromArgb(255, 255, 182, 193); // LightPink

                // Change the Accent6 theme color to the pastel shade
                workbook.SetThemeColor(ThemeColorType.Accent6, pastelAccent6);

                // Verify chart legends in all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each chart on the worksheet
                    for (int i = 0; i < sheet.Charts.Count; i++)
                    {
                        try
                        {
                            Chart chart = sheet.Charts[i];
                            Legend legend = chart.Legend;

                            // Output legend verification details
                            Console.WriteLine($"Worksheet: {sheet.Name}, Chart Index: {i}");
                            Console.WriteLine($"  Legend Position: {legend.Position}");
                            // Note: Aspose.Cells Legend does not expose an IsVisible property in current versions.
                        }
                        catch (Exception exChart)
                        {
                            Console.WriteLine($"Error processing chart {i} on sheet \"{sheet.Name}\": {exChart.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
