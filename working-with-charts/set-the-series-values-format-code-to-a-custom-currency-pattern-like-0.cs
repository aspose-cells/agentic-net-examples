// Title: Apply Custom Currency Format ($#,##0) to Chart Series Values with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, bind the data range, and set the series ValuesFormatCode to the custom pattern "$#,##0" so the chart displays values as currency (e.g., $1,234) before saving the file.
// Keywords: Aspose.Cells chart format code | ValuesFormatCode C# | custom currency format Aspose.Cells | chart series number formatting .NET | Excel chart currency display
// Common Searches: Aspose.Cells set chart series format code | custom number format for chart values C# | ValuesFormatCode example Aspose.Cells | format chart data as currency .NET | apply $#,##0 pattern to Excel chart series
// Developer Intent: Apply a custom currency number format to the values of a chart series using Aspose.Cells for .NET.
// Use Cases: Display sales figures in a column chart with dollar signs and thousand separators. | Build a financial dashboard where chart data points are shown as formatted currency. | Generate Excel reports with charts that use the "$#,##0" pattern for all series values.
// AI Prompts: Show a C# code snippet that sets ValuesFormatCode to "$#,##0" for a chart series in Aspose.Cells. | Explain step‑by‑step how to apply a custom number format to chart series values with Aspose.Cells for .NET. | Provide guidance on formatting Excel chart data labels using a custom currency pattern in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, insert a column chart, bind the data range, and set the series ValuesFormatCode to the custom pattern "$#,##0" so the chart displays values as currency (e.g., $1,234) before saving the file.
    public class SeriesValuesFormatCodeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(1234);
                worksheet.Cells["B3"].PutValue(5678);
                worksheet.Cells["B4"].PutValue(9012);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set custom currency format for the series values
                // "$#,##0" will display numbers like $1,234
                chart.NSeries[0].ValuesFormatCode = "$#,##0";

                // Save the workbook
                workbook.Save("SeriesValuesFormatCodeDemo_out.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SeriesValuesFormatCodeDemo.Run();
        }
    }
}
