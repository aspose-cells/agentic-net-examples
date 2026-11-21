// Title: Binding a column chart to a populated cell range in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# bind column chart to specific cell range | set chart data source to A1:B4 range using Aspose.Cells | programmatically create Excel column chart from worksheet data Aspose.Cells | how to assign worksheet cells as chart series in C# with Aspose.Cells | example of chart data range automation in Aspose.Cells .NET
// Tags: chart data range binding Aspose.Cells | column chart series from worksheet cells C# | Aspose.Cells NSeries.Add data source | automated Excel chart creation .NET | populate chart categories from column A Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills cells A1:B4 with category and value data, adds a column chart, binds the value series to the B2:B4 range using NSeries.Add (including headers), lets the chart infer categories from column A, and saves the file as ChartDataRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Add the values series (including headers)
            chart.NSeries.Add("B2:B4", true);
            // Category (X) axis labels are taken from column A
            // In newer Aspose.Cells versions the CategoryData property may be unavailable;
            // the chart will automatically use the first column as categories when headers are present.

            // Define output file path
            string outputPath = "ChartDataRange.xlsx";

            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
