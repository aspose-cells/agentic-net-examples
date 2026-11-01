// Title: Create a line chart from range A1:A5 and save the workbook using Aspose.Cells for .NET
// Common Searches: how to add a line chart from a cell range using Aspose.Cells C# | Aspose.Cells example to set chart title programmatically in .NET | save an Excel workbook to a specific folder with Aspose.Cells C# | generate an Excel file with a line chart using Aspose.Cells | add a chart to a worksheet at a custom position with Aspose.Cells
// Tags: line chart creation Aspose.Cells C# | add chart series from range Aspose.Cells | set chart title Aspose.Cells .NET | save workbook to file Aspose.Cells | ensure output directory exists C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart and ChartType enums

// The sample creates a new workbook, writes numeric values to cells A1 through A5, adds a line chart covering that range, sets the chart title, ensures the output directory exists, and saves the workbook as SampleChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in A1:A5
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(-5);
            sheet.Cells["A3"].PutValue(15);
            sheet.Cells["A4"].PutValue(-8);
            sheet.Cells["A5"].PutValue(20);

            // Add a line chart as an alternative to sparkline
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A5", true);
            chart.Title.Text = "Sample Line Chart";

            // Define output file path
            string outputPath = "SampleChart.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
