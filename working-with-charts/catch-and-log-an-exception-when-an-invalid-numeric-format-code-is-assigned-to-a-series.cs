// Title: Handle CellsException for invalid chart series numeric format in Aspose.Cells .NET
// Description: Creates a workbook, adds a column chart, enables strict custom number‑format checking, attempts to assign an invalid ValuesFormatCode to the first series, catches the resulting CellsException, logs its message and code, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | chart series | ValuesFormatCode | invalid numeric format | CellsException | CheckCustomNumberFormat | error handling | Excel chart
// Common Searches: Aspose.Cells catch CellsException | invalid ValuesFormatCode error | CheckCustomNumberFormat true example | chart series numeric format exception .NET | log Aspose.Cells exception
// Developer Intent: Show how to detect and log the exception thrown when an unsupported numeric format is applied to a chart series.
// Use Cases: Automated Excel report generation where user‑supplied number formats may be invalid. | Validating custom number formats before applying them to chart series. | Logging detailed exception information for troubleshooting workbook creation scripts. | Implementing resilient error handling in Excel automation pipelines.
// AI Prompts: Generate C# code that validates a format string before setting ValuesFormatCode and logs any CellsException. | Explain how workbook.Settings.CheckCustomNumberFormat influences exception throwing for invalid chart format codes. | Provide a mapping from CellsException.Code to user‑friendly error messages for chart formatting failures. | Suggest a retry or fallback strategy after catching an invalid numeric format exception in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, enables strict custom number‑format checking, attempts to assign an invalid ValuesFormatCode to the first series, catches the resulting CellsException, logs its message and code, and saves the file.
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

        // Enable strict checking of custom number formats
        workbook.Settings.CheckCustomNumberFormat = true;

        try
        {
            // Attempt to assign an invalid numeric format code to the series
            // This should trigger a CellsException because the format is not valid
            chart.NSeries[0].ValuesFormatCode = "invalid_format_code";
        }
        catch (CellsException ex)
        {
            // Log exception details
            Console.WriteLine("Caught exception while setting ValuesFormatCode:");
            Console.WriteLine("Message: " + ex.Message);
            Console.WriteLine("Exception Type: " + ex.Code);
        }

        // Save the workbook
        workbook.Save("SeriesInvalidFormatDemo_out.xlsx");
    }
}
