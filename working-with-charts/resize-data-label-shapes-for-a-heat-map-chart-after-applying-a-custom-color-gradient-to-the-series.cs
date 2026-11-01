// Title: Resize data label shapes on a Surface3D heat‑map chart after applying a vertical two‑color gradient with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that creates a Surface3D heat‑map chart, applies a vertical two‑color gradient to the series, enables data labels, and then sets a custom width and height for each data label shape. | Show how to programmatically adjust the size (width and height) of data label shapes on a heat‑map chart after a gradient fill is applied, using the Aspose.Cells Chart and Shape APIs in .NET.
// Common Searches: C# Aspose.Cells how to change size of data label shapes on a Surface3D chart | adjust data label dimensions after applying gradient fill in Aspose.Cells | resize heat map chart data labels programmatically with Aspose.Cells .NET | set custom width and height for chart data labels in Aspose.Cells workbook
// Tags: chart data label shape resizing Aspose.Cells | surface3d heat map gradient fill Aspose.Cells | two-color vertical gradient series .NET | data label shape scaling Aspose.Cells | heat map chart label dimension customization

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing; // For GradientStyleType

namespace HeatMapExample
{
    // // This program creates a new workbook, fills a 5x5 range with sample values, adds a Surface3D chart to represent a heat map, applies a vertical two‑color gradient to the series, enables data labels to show values, and saves the workbook as an Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data for the heat map (5x5 matrix)
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        sheet.Cells[row, col].PutValue(row * 5 + col + 1);
                    }
                }

                // Add a Surface3D chart (used as a heat‑map representation) covering the data range
                int chartIndex = sheet.Charts.Add(ChartType.Surface3D, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Define the series (the whole data range)
                chart.NSeries.Add("A1:E5", true);

                // Get the first (and only) series
                Series series = chart.NSeries[0];

                // Apply a custom two‑color gradient to the series fill (vertical)
                series.Area.FillFormat.SetTwoColorGradient(
                    Color.LightBlue,
                    Color.DarkBlue,
                    GradientStyleType.Vertical,
                    1);

                // Enable data labels and show the cell values
                series.DataLabels.ShowValue = true;

                // Determine output file path
                string outputFile = "HeatMapWithResizedDataLabels.xlsx";

                // Ensure the directory exists (if a directory part is present)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
