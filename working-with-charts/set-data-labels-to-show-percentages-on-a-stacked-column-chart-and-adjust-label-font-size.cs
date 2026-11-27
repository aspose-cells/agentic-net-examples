// Title: Add percentage data labels with a specific font size to a stacked column chart using Aspose.Cells for .NET
// Common Searches: how to display percentage labels on a stacked column chart with Aspose.Cells in C# | set font size for chart data labels in Aspose.Cells .NET | configure data label position and hide values in Aspose.Cells stacked column chart | Aspose.Cells show only percentages in stacked column chart labels | change chart label color and size for Excel chart using Aspose.Cells
// Tags: stacked column chart show percentages Aspose.Cells | chart label styling Aspose.Cells | exclude raw values from chart labels Aspose.Cells | inside end label placement Aspose.Cells | set chart label color Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a stacked column chart, and configures each series' data labels to show only percentages, positions them InsideEnd, sets a custom font size and color, then saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a stacked column chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Series1");
            worksheet.Cells["C1"].PutValue("Series2");
            worksheet.Cells["D1"].PutValue("Series3");

            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            worksheet.Cells["C2"].PutValue(15);
            worksheet.Cells["C3"].PutValue(25);
            worksheet.Cells["C4"].PutValue(35);

            worksheet.Cells["D2"].PutValue(5);
            worksheet.Cells["D3"].PutValue(10);
            worksheet.Cells["D4"].PutValue(15);

            // Add a stacked column chart to the worksheet
            // Correct chart type enum is ColumnStacked
            int chartIndex = worksheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Add three series (one for each data column)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries.Add("D2:D4", true);

            // Set the common category (X‑axis) data for all series
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels for each series to show percentages and set font size
            foreach (Series series in chart.NSeries)
            {
                DataLabels dataLabels = series.DataLabels;
                dataLabels.ShowPercentage = true;               // Show percentages
                dataLabels.ShowValue = false;                   // Hide raw values (optional)
                dataLabels.Position = LabelPositionType.InsideEnd; // Good position for stacked columns
                dataLabels.Font.Size = 10;                      // Adjust font size
                dataLabels.Font.Color = Color.Black;            // Optional: set font color
            }

            // Save the workbook
            string outputPath = "StackedColumnChart_PercentLabels.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
