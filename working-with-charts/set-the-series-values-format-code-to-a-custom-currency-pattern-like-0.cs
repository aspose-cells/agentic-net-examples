// Title: How to apply a custom currency format ($#,##0) to Excel chart series values using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set ValuesFormatCode for chart series to custom currency pattern | C# Aspose.Cells column chart series custom number format $#,##0 | Apply custom currency format to Excel chart data series with Aspose.Cells .NET | Set custom number format for chart series values in Aspose.Cells workbook | Aspose.Cells example for formatting chart series as currency in C#
// Tags: Aspose.Cells series format code | C# Excel chart series number formatting | Aspose.Cells column chart custom formatting | Aspose.Cells currency format code $#,##0 | format chart series values .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SeriesCustomCurrencyFormatDemo
{
    // The example creates a workbook, adds category and amount data, inserts a column chart, assigns the amount range as the series values, and uses the ValuesFormatCode property to apply the custom currency pattern "$#,##0" to the series before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                // Column A: Categories
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                // Column B: Numeric values (currency amounts)
                worksheet.Cells["B1"].PutValue("Amount");
                worksheet.Cells["B2"].PutValue(1234.56);
                worksheet.Cells["B3"].PutValue(7890.12);
                worksheet.Cells["B4"].PutValue(345.67);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.Title.Text = "Sample Currency Chart";

                // Set the data source for the series (values) and categories
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";     // Categories

                // Apply a custom currency format to the first series values
                // The format code "$#,##0" will display numbers like $1,235
                Series firstSeries = chart.NSeries[0];
                firstSeries.ValuesFormatCode = "$#,##0";

                // Determine output path and ensure directory exists
                string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                string outputPath = Path.Combine(outputDir, "SeriesCustomCurrencyFormat.xlsx");

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors
                Console.Error.WriteLine("An error occurred while creating the workbook:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
