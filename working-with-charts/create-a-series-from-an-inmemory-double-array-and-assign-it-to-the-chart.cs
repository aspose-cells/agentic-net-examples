// Title: Create a column chart in an Excel workbook from an in‑memory double array using Aspose.Cells for .NET
// AI Prompts: Write C# code that populates a worksheet with a double[] and adds a column chart whose series references the written cells using Aspose.Cells. | Show how to build an Aspose.Cells NSeries from a range that was filled from an in‑memory numeric array without intermediate files. | Provide a complete example that saves the workbook as an .xlsx file after creating the chart from the double array data.
// Common Searches: Aspose.Cells C# create column chart from double array without CSV | how to bind in‑memory numeric array to Excel chart series using Aspose.Cells | add chart series from worksheet range populated by double[] in .NET | generate Excel file with chart from array data using Aspose.Cells for .NET
// Tags: add column chart from double array Aspose.Cells | populate worksheet cells with in‑memory array C# | define NSeries using cell range Aspose.Cells | save workbook with chart as .xlsx Aspose.Cells | chart data source from in‑memory numeric array .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, writes a double[] into column A, adds a column chart, sets the series to the A1:A4 range, names the series, and saves the file as ChartFromArray.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Initialize a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // In‑memory double array
            double[] values = new double[] { 10.5, 20.0, 30.75, 40.25 };

            // Write the array values to cells (column A)
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]);
            }

            // Add a Column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Create a series for the chart using the range that contains the array data
            int seriesIndex = chart.NSeries.Add($"A1:A{values.Length}", true);
            chart.NSeries[seriesIndex].Name = "Sample Data";

            // Define output file path
            string outputPath = "ChartFromArray.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
