// Title: How to catch and log an exception when assigning an invalid NumberFormat to a chart series data label in Aspose.Cells for .NET
// AI Prompts: Write C# code that wraps chartSeries.DataLabels.NumberFormat assignment in a try/catch block and writes the exception message to the console using Aspose.Cells. | Show how to validate a numeric format string before setting it on a chart series data label and log any errors in a .NET application. | Generate a complete Aspose.Cells example that creates a column chart, attempts to set an invalid NumberFormat on its series, and captures the exception for logging.
// Common Searches: Aspose.Cells set data label number format exception handling C# | C# catch error when using invalid NumberFormat on chart series Aspose.Cells | how to log exception for wrong numeric format code in Aspose.Cells chart
// Tags: chart series data label number format exception handling | Aspose.Cells numeric format validation | C# try catch for chart data label formatting | column chart series format error logging Aspose.Cells | invalid NumberFormat code handling in .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// // Demonstrates creating a workbook, adding a column chart with a series, attempting to assign an invalid NumberFormat to the series' data labels, catching and logging the resulting exception, and saving the workbook to an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);

            // Add a column chart to the worksheet (rows 5-15, columns 0-5)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIdx];

            // Add a series to the chart using data from column A
            int seriesIdx = chart.NSeries.Add("A1:A2", true);

            // Attempt to assign an invalid numeric format code to the series' data labels
            try
            {
                // This string is not a valid numeric format and will trigger an exception
                chart.NSeries[seriesIdx].DataLabels.NumberFormat = "InvalidFormatCode";
            }
            catch (Exception ex)
            {
                // Log the exception details
                Console.WriteLine("Exception caught while setting numeric format: " + ex.Message);
            }

            // Save the workbook to a file
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully as output.xlsx");
        }
        catch (Exception ex)
        {
            // Log any unexpected exceptions
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
