// Title: How to catch and log a CellsException when assigning an invalid ValuesFormatCode to a chart series in Aspose.Cells for .NET
// AI Prompts: Generate C# code that wraps the assignment of ValuesFormatCode on an Aspose.Cells chart series in a try‑catch block and writes the CellsException Code and Message to the console. | Show how to log both Aspose.Cells specific CellsException and a generic Exception after a failed ValuesFormatCode assignment, then continue to save the workbook. | Provide a reusable method that receives a chart series and a format string, attempts to set ValuesFormatCode, catches any CellsException, logs the details, and returns a success flag.
// Common Searches: aspnet catch CellsException when setting chart series numeric format code | aspose.cells invalid ValuesFormatCode exception handling example | c# try catch for chart series format code error in Aspose.Cells | log error code from CellsException after assigning wrong format to chart series
// Tags: handle CellsException for chart series formatting | log Aspose.Cells format code errors | validate chart series ValuesFormatCode .NET | exception handling during chart creation Aspose.Cells | save workbook after chart formatting failure

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, attempts to set an invalid ValuesFormatCode on the first series, and demonstrates try‑catch handling that logs CellsException details and any other unexpected errors before saving the file.
class SeriesInvalidFormatDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(0.25);
        worksheet.Cells["B3"].PutValue(0.5);
        worksheet.Cells["B4"].PutValue(0.75);

        // Add a column chart and set its data source
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Try to assign an invalid numeric format code to the series
        try
        {
            // Deliberately incorrect format string
            chart.NSeries[0].ValuesFormatCode = "invalid_format@@@";
            Console.WriteLine("Format code assigned without error.");
        }
        catch (CellsException ex)
        {
            // Log Aspose.Cells specific exception details
            Console.WriteLine("Error assigning format code to series:");
            Console.WriteLine("Exception Type: " + ex.Code);
            Console.WriteLine("Message: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Log any other unexpected exceptions
            Console.WriteLine("Unexpected error: " + ex.Message);
        }

        // Save the workbook
        workbook.Save("SeriesInvalidFormatDemo_out.xlsx");
    }
}
