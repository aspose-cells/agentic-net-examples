// Title: C# – Set Accent6 Theme Color to Pastel Pink and Ensure Chart Legends Are Visible with Aspose.Cells
// Description: Loads an existing Excel workbook, changes the Accent6 theme color to a pastel pink (RGB 255,182,193), iterates through every worksheet and chart to guarantee that each chart legend is shown, logs the verification details, and saves the updated file as a new workbook.
// Keywords: Aspose.Cells C# set Accent6 theme color | pastel pink theme Excel Aspose | chart legend visibility Aspose.Cells | modify Excel theme programmatically .NET | ensure chart legends show | theme color Accent6 pastel | Aspose.Cells workbook styling US | Excel automation C# Aspose
// Common Searches: How to change Accent6 theme color in Aspose.Cells for .NET | Set custom pastel theme color in Excel using C# | Make all chart legends visible with Aspose.Cells | Aspose.Cells example for updating theme colors and chart legends | C# code to verify chart legend visibility in Excel
// Developer Intent: Load a workbook, apply a pastel pink shade to the Accent6 theme color, confirm that every chart’s legend is enabled, and save the modified workbook.
// Use Cases: Standardize corporate reports with a pastel color palette while keeping chart legends readable. | Prepare presentation‑ready Excel files by programmatically applying a custom theme and ensuring legends are displayed. | Batch‑process multiple workbooks to enforce a consistent Accent6 color and legend visibility across all charts.
// AI Prompts: Write C# code using Aspose.Cells to set the Accent6 theme color to RGB(255,182,193) and save the workbook. | Create a method that loops through all worksheets and charts in an Excel file, sets ShowLegend = true for each chart, and logs worksheet name, chart index, and legend status.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsThemeAndChartDemo
{
    // Loads an existing Excel workbook, changes the Accent6 theme color to a pastel pink (RGB 255,182,193), iterates through every worksheet and chart to guarantee that each chart legend is shown, logs the verification details, and saves the updated file as a new workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the input workbook
                string inputPath = "input.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Change the Accent6 theme color to a pastel shade (Light Pink)
                Color pastelAccent6 = Color.FromArgb(255, 255, 182, 193);
                workbook.SetThemeColor(ThemeColorType.Accent6, pastelAccent6);

                // Verify and ensure chart legends are visible in each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Ensure the legend is visible; if not, make it visible
                        if (!chart.ShowLegend)
                        {
                            chart.ShowLegend = true;
                        }

                        // Output legend verification info
                        Console.WriteLine($"Worksheet: {sheet.Name}, Chart Index: {sheet.Charts.IndexOf(chart)}");
                        Console.WriteLine($"  Legend Visible: {chart.ShowLegend}");
                    }
                }

                // Path to the output workbook
                string outputPath = "output.xlsx";

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine("Workbook saved with updated Accent6 theme color and verified chart legends.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
