using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesFormatExceptionDemo
{
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

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Attempt to assign an invalid numeric format code to the series
            try
            {
                // This format string is deliberately invalid and should trigger an exception
                chart.NSeries[0].ValuesFormatCode = "invalid_format_code";
                Console.WriteLine("Format code assigned successfully (unexpected).");
            }
            catch (CellsException ex)
            {
                // Log the exception details
                Console.WriteLine("Caught CellsException while setting ValuesFormatCode:");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"Exception Type Code: {ex.Code}");
            }
            catch (Exception ex)
            {
                // Catch any other unexpected exceptions
                Console.WriteLine("Caught unexpected exception:");
                Console.WriteLine($"Message: {ex.Message}");
            }

            // Save the workbook (even if the format code was invalid, the workbook can still be saved)
            workbook.Save("SeriesFormatExceptionDemo_out.xlsx");
        }
    }
}