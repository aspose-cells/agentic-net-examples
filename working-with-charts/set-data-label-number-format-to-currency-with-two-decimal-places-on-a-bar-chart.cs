// Title: Format Bar Chart Data Labels as Currency (Two Decimals) with Aspose.Cells for .NET
// Description: This example creates a workbook, adds category and value data, inserts a bar chart, enables data labels, and applies the custom number format "$#,##0.00" so each label appears as currency with two decimal places before saving the file.
// Keywords: Aspose.Cells C# chart formatting | bar chart data label currency | number format Excel chart .NET | custom number format Aspose.Cells | currency display on chart labels | Excel automation financial charts | C# Aspose.Cells example | chart series label formatting
// Common Searches: Aspose.Cells set currency format for chart labels | C# bar chart data label number format | apply custom number format to Excel chart series | format chart data labels as money using Aspose.Cells | how to show $ values on bar chart labels .NET
// Developer Intent: Apply a monetary number format with two decimal places to the data labels of a bar chart in a .NET workbook.
// Use Cases: Generate a sales performance chart where each bar shows the amount as $1,234.56. | Automate quarterly revenue reports with bar charts that display monetary values on the labels. | Create financial dashboards in Excel that require currency‑styled data labels for easy reading.
// AI Prompts: Show C# code to set a custom currency format for bar chart data labels using Aspose.Cells. | How can I display monetary values with two decimal places on chart labels in a .NET workbook? | Explain how to assign different number formats to multiple series' data labels in an Aspose.Cells chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds category and value data, inserts a bar chart, enables data labels, and applies the custom number format "$#,##0.00" so each label appears as currency with two decimal places before saving the file.
    public class BarChartDataLabelCurrencyFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the bar chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(1000);
                worksheet.Cells["B3"].PutValue(2000);
                worksheet.Cells["B4"].PutValue(3000);

                // Add a bar chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Bar, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Set data label number format to currency with two decimal places
                series.DataLabels.NumberFormat = "$#,##0.00";

                // Define output file path
                string outputPath = "BarChartDataLabelCurrency.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            BarChartDataLabelCurrencyFormat.Run();
        }
    }
}
