// Title: Apply Currency Number Format to Waterfall Chart Axis with Aspose.Cells for .NET
// Description: Shows how to build a workbook, insert a Waterfall chart, and style the value‑axis tick labels with a dollar format ($#,##0) via the TickLabels.NumberFormat property in C#.
// Keywords: Aspose.Cells | C# chart formatting | waterfall chart axis | currency number format | TickLabels.NumberFormat | Excel chart automation | set chart axis format .NET | financial waterfall chart
// Common Searches: Aspose.Cells set axis number format | waterfall chart currency format C# | format chart value axis as $#,##0 | change tick label format Aspose.Cells | apply number format to Excel chart axis .NET
// Developer Intent: Format the value axis of a Waterfall chart as currency.
// Use Cases: Generate financial waterfall charts where axis values display dollar amounts. | Automate Excel reports with correctly formatted chart axes for accounting dashboards. | Create reusable code that enforces consistent currency display across multiple charts.
// AI Prompts: How do I set a currency number format for a Waterfall chart axis using Aspose.Cells in C#? | Provide a C# example that applies $#,##0 format to chart tick labels with Aspose.Cells. | Explain the steps to use TickLabels.NumberFormat for chart axis formatting in Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, insert a Waterfall chart, and style the value‑axis tick labels with a dollar format ($#,##0) via the TickLabels.NumberFormat property in C#.
class WaterfallChartNumberFormat
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for the waterfall chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Start");
            worksheet.Cells["A3"].PutValue("Increase");
            worksheet.Cells["A4"].PutValue("Decrease");
            worksheet.Cells["A5"].PutValue("End");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(5000);
            worksheet.Cells["B3"].PutValue(2000);
            worksheet.Cells["B4"].PutValue(-1500);
            worksheet.Cells["B5"].PutValue(5500);

            // Add a Waterfall chart
            int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Set custom number format for the value axis tick labels to display currency
            chart.ValueAxis.TickLabels.NumberFormat = "$#,##0";

            // Save the workbook
            string outputPath = "WaterfallChartCurrencyFormat.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        WaterfallChartNumberFormat.Run();
    }
}
