// Title: Generate a column chart in an Excel workbook from an in‑memory double array using Aspose.Cells for .NET
// Common Searches: how to bind a double[] to an Excel column chart with Aspose.Cells C# | Aspose.Cells create chart series from in‑memory numeric array | C# add column chart series using NSeries.Add range from array values | populate Excel chart with values written to worksheet cells using Aspose.Cells | save workbook with chart sourced from double array Aspose.Cells .NET
// Tags: create column chart series Aspose.Cells | bind in‑memory double array to Excel chart | NSeries.Add range Aspose.Cells C# | write numeric array to worksheet cells Aspose.Cells | save workbook with chart Aspose.Cells .xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes a double[] to column A of the first worksheet, adds a column chart, links the array values to the chart series via a range reference, sets a chart title, and saves the workbook as an .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // In‑memory double array that will be used as series values
            double[] values = new double[] { 10.5, 20.0, 30.2, 40.8, 50.1 };

            // Write the array values to the worksheet (column A)
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]);
            }

            // Add a column chart to the worksheet (positioned from row 5, column 0 to row 20, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the series to the chart using the range that contains the values
            string range = $"{sheet.Name}!A1:A{values.Length}";
            chart.NSeries.Add(range, true);

            // Optional: set a title for the chart
            chart.Title.Text = "Sample Chart from Array";

            // Define output file path
            string outputPath = "ChartFromArray.xlsx";

            // Ensure the output directory exists
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
