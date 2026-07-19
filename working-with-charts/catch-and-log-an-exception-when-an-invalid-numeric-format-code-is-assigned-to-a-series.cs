// Title: Catch CellsException for an invalid ValuesFormatCode on a chart series – Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a column chart, and deliberately assigns an invalid numeric format string to the first series. The assignment is wrapped in a try‑catch block that captures the CellsException, logs its message and error code, and then saves the workbook.
// Keywords: Aspose.Cells | .NET | C# | chart series | ValuesFormatCode | invalid format code | CellsException | exception handling | logging | SeriesInvalidFormatCodeDemo | Aspose.Cells.Charts | numeric format error
// Common Searches: Aspose.Cells catch CellsException when setting ValuesFormatCode | invalid numeric format code chart series Aspose.Cells | how to log exception for chart series format error in .NET | example of try‑catch around ValuesFormatCode assignment | Aspose.Cells chart series format code validation
// Developer Intent: Capture and log the CellsException thrown by assigning an invalid ValuesFormatCode to a chart series in Aspose.Cells.
// Use Cases: Validate a format string before assigning it to chart.NSeries[i].ValuesFormatCode to avoid runtime errors. | Wrap the ValuesFormatCode assignment in a try‑catch block, log the exception details, and continue workbook processing. | Apply a fallback numeric format when a CellsException occurs, ensuring the chart remains displayable.
// AI Prompts: Generate a utility method that checks a numeric format string for validity and safely assigns it to a chart series, logging any CellsException. | Show how to implement a default format fallback inside a catch block when setting ValuesFormatCode fails. | Write code that records CellsException details to a log file instead of the console for invalid format codes in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds a column chart, and deliberately assigns an invalid numeric format string to the first series. The assignment is wrapped in a try‑catch block that captures the CellsException, logs its message and error code, and then saves the workbook.
    public class SeriesInvalidFormatCodeDemo
    {
        public static void Run()
        {
            try
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
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Attempt to assign an invalid numeric format code to the first series
                try
                {
                    // This format string is deliberately invalid and should trigger an exception
                    chart.NSeries[0].ValuesFormatCode = "invalid_format_code";
                }
                catch (CellsException ex)
                {
                    // Log the exception details
                    Console.WriteLine("Exception caught while setting ValuesFormatCode:");
                    Console.WriteLine("Message: " + ex.Message);
                    Console.WriteLine("Exception Type Code: " + ex.Code);
                }

                // Save the workbook
                workbook.Save("SeriesInvalidFormatCodeDemo_out.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SeriesInvalidFormatCodeDemo.Run();
        }
    }
}
