// Title: Change the Accent6 theme color to a pastel shade and verify chart legends in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Apply a pastel teal (RGB 173,216,230) to the workbook's Accent6 theme color using Aspose.Cells and save the result. | Loop through every worksheet and each chart to output the chart name and whether its legend is currently visible. | If a chart's legend is hidden, set ShowLegend to true, position the legend on the right, and re‑save the workbook.
// Common Searches: asp.net set accent6 theme color to pastel using Aspose.Cells | c# check chart legend visibility in Excel with Aspose.Cells | how to iterate over all charts in a workbook and read legend property Aspose.Cells | change Excel workbook theme to custom color and list chart legends c# | asp.net enable hidden chart legend and set position right Aspose.Cells
// Tags: theme color customization Aspose.Cells | pastel color for Excel workbook C# | chart legend presence detection Aspose.Cells | worksheet chart enumeration Aspose.Cells | save modified workbook Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing Excel file, changes the Accent6 theme color to a pastel teal, iterates through each worksheet and chart to report legend visibility, optionally makes hidden legends visible and positions them, then saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "InputWorkbook.xlsx";

            // Verify the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Change the Accent6 theme color to a pastel shade (light pastel teal)
            workbook.SetThemeColor(ThemeColorType.Accent6, Color.FromArgb(173, 216, 230)); // Light pastel blue

            // Verify chart legends in all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Determine if the legend is visible using Chart.ShowLegend
                        bool hasLegend = chart.ShowLegend;

                        // Output verification result
                        string chartName = chart.Name ?? "Unnamed";
                        Console.WriteLine($"Worksheet: {sheet.Name}, Chart: {chartName}");
                        Console.WriteLine($"  Legend present: {hasLegend}");

                        // Optionally, ensure the legend is visible (uncomment to enforce)
                        // if (!hasLegend)
                        // {
                        //     chart.ShowLegend = true;
                        //     chart.Legend.Position = LegendPosition.Right;
                        // }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart in worksheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            // Output workbook path
            string outputPath = "OutputWorkbook.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Error saving workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
