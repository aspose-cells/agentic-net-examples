// Title: How to bind data label number format to source cells for the first series in an Aspose.Cells column chart (C#)
// Common Searches: Aspose.Cells C# set DataLabels.NumberFormatLinked for column chart series | link data label number format to worksheet cells in Aspose.Cells chart | enable source cell number format on data labels of first series using Aspose.Cells | C# Aspose.Cells chart data labels inherit cell formatting | bind data label formatting to source range in Aspose.Cells column chart
// Tags: DataLabels.NumberFormatLinked property Aspose.Cells | column chart first series data label formatting C# | bind data label format to source cells Aspose.Cells | Aspose.Cells chart data label number format linking | C# Aspose.Cells chart series formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, adds the first series, enables data labels, sets DataLabels.NumberFormatLinked to true so the labels inherit the number format from the source cells, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Get the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Add the first series (values). The second argument (true) indicates that the series is vertical.
            chart.NSeries.Add("B2:B4", true);

            // Enable data labels for the first series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Bind the number format of data labels to the source cell format
            chart.NSeries[0].DataLabels.NumberFormatLinked = true;

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
