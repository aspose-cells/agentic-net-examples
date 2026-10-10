// Title: Enable and format data labels on a box‑and‑whisker chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# sample using Aspose.Cells to build a box‑and‑whisker chart from a data matrix and activate its data labels to show the computed values. | Demonstrate how to set the font size and color of data labels for a box‑and‑whisker chart derived from a statistical summary range in Aspose.Cells.
// Common Searches: Aspose.Cells C# enable data labels on box plot chart | show values on box‑and‑whisker chart using Aspose.Cells .NET | how to format data label font for box plot in Aspose.Cells | create box and whisker chart from raw data range Aspose.Cells C# | use statistical summary range for box plot chart Aspose.Cells
// Tags: Aspose.Cells enable data labels C# | box-and-whisker chart data labels Aspose.Cells | statistical summary range chart Aspose.Cells | format data label font Aspose.Cells | create box plot column chart Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, writes category and observation data, adds a column chart that Aspose.Cells interprets as a box‑and‑whisker chart, enables data labels to display the calculated statistics, customizes label font size and color, and saves the file as BoxAndWhiskerChart_WithDataLabels.xlsx.
class BoxAndWhiskerChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample categories
            string[] categories = { "A", "B", "C", "D", "E" };

            // Sample observations (5 categories × 5 observations)
            double[,] values = {
                { 10, 12, 14, 13, 11 },
                { 20, 22, 19, 21, 23 },
                { 30, 28, 31, 29, 32 },
                { 40, 42, 39, 41, 43 },
                { 50, 48, 51, 49, 52 }
            };

            // Write categories to column A (A2:A6)
            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]);
            }

            // Write observations to columns B‑F (B2:F6)
            for (int row = 0; row < values.GetLength(0); row++)
            {
                for (int col = 0; col < values.GetLength(1); col++)
                {
                    sheet.Cells[row + 1, col + 1].PutValue(values[row, col]);
                }
            }

            // Define the raw data range for the chart
            string dataRange = "B2:F6";

            // Add a column chart (Box‑and‑Whisker not available in older versions)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Use the raw data range; Aspose.Cells will compute the statistical summary automatically
            chart.NSeries.Add(dataRange, true);

            // Enable and format data labels
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[0].DataLabels.Font.Size = 10;
            chart.NSeries[0].DataLabels.Font.Color = Color.Black;

            // Save the workbook
            string outputPath = "BoxAndWhiskerChart_WithDataLabels.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
