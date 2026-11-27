// Title: How to catch and log a CellsException when setting an invalid ValuesFormatCode on a chart series in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells catch CellsException for invalid chart series format code | log error when ValuesFormatCode is wrong in Aspose.Cells chart | handle exception while assigning numeric format string to chart series using Aspose.Cells | Aspose.Cells chart series invalid format code exception handling example | write CellsException details to file in C# Aspose.Cells
// Tags: Aspose.Cells chart series format code exception | catch CellsException in .NET | log Excel chart formatting errors Aspose.Cells | invalid ValuesFormatCode handling Aspose.Cells | C# exception handling Aspose.Cells chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data and a column chart, then deliberately assigns an invalid ValuesFormatCode to the first series. The resulting CellsException is caught, its message is printed to the console, and the full exception details are appended to an error_log.txt file before the workbook is saved.
class Program
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

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Attempt to assign an invalid numeric format code to the series
        try
        {
            // This format string is deliberately invalid and should trigger an exception
            chart.NSeries[0].ValuesFormatCode = "invalid@@format";
        }
        catch (CellsException ex)
        {
            // Log the exception details to the console
            Console.WriteLine("Caught CellsException while setting ValuesFormatCode: " + ex.Message);

            // Additionally, write the full exception stack trace to a log file
            string logFile = "error_log.txt";
            string logEntry = $"{DateTime.UtcNow:O} - {ex.GetType().FullName}: {ex.Message}{Environment.NewLine}{ex.StackTrace}{Environment.NewLine}";
            File.AppendAllText(logFile, logEntry);
        }

        // Save the workbook to verify that the rest of the operations succeeded
        string outputFile = "SeriesInvalidFormatDemo.xlsx";
        workbook.Save(outputFile);
        Console.WriteLine($"Workbook saved as '{outputFile}'.");
        Console.WriteLine("If an invalid format was supplied, details are logged in 'error_log.txt'.");
    }
}
