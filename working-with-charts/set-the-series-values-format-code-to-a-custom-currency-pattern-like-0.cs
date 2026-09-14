// Title: Set a custom currency number format for a chart series using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells and applies a custom currency pattern (e.g., "$#,##0") to the series values via the ValuesFormatCode property. | Show how to modify the currency symbol in a chart series number format (for example, changing to "€#,##0.00") using Aspose.Cells in a .NET application. | Provide a snippet that opens an existing workbook, updates the ValuesFormatCode of a specific chart series to a custom format, and saves the workbook.
// Common Searches: Aspose.Cells C# set custom currency format for chart series values | how to use ValuesFormatCode to format chart data as currency in Aspose.Cells .NET | example of applying "$#,##0" number format to Excel chart series with Aspose.Cells | change currency symbol for chart series values in Aspose.Cells C# code | formatting column chart series as currency using Aspose.Cells for .NET
// Tags: chart series custom currency format Aspose.Cells | ValuesFormatCode property usage .NET | apply number format to Excel chart series C# | Aspose.Cells column chart formatting example | set series values format code C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding sample data, inserting a column chart, and using the ValuesFormatCode property to apply a custom currency pattern like "$#,##0" to the chart series values before saving.
    public class SetSeriesValuesCustomCurrencyFormat
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
                worksheet.Cells["B2"].PutValue(1250);
                worksheet.Cells["B3"].PutValue(2500);
                worksheet.Cells["B4"].PutValue(3750);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply a custom currency format to the series values
                // "$#,##0" will display numbers like $1,250
                chart.NSeries[0].ValuesFormatCode = "$#,##0";

                // Save the workbook
                string outputPath = "SeriesValuesCustomCurrencyFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetSeriesValuesCustomCurrencyFormat.Run();
        }
    }
}
