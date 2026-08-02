// Title: Format Bar Chart Data Labels as Currency (Two Decimals) with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a 2‑D bar chart, enable data labels, and apply a currency number format with two decimal places ("$#,##0.00" and built‑in format index 2) to the labels using Aspose.Cells for C#. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells | C# chart formatting | bar chart data labels | currency number format | two decimal places | NumberFormat property | built‑in format index | Excel export | financial chart
// Common Searches: Aspose.Cells format chart data labels as currency | C# set number format for bar chart labels | currency format with two decimals in Aspose.Cells chart | apply built‑in currency format index to chart labels | how to show dollar values on Excel bar chart using Aspose
// Developer Intent: Apply a currency number format with two decimal places to the data labels of a bar chart in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate financial dashboards where bar chart labels display amounts in dollars and cents. | Create sales reports with bar charts that automatically format values as currency for readability. | Export analytics to Excel while ensuring chart labels follow locale‑specific currency formatting.
// AI Prompts: Write C# code with Aspose.Cells that adds a bar chart and formats its data labels as currency with two decimal places. | Show how to set both a custom number format string and the built‑in currency format index for chart data labels in Aspose.Cells. | Explain how to modify the number format of existing chart data labels after the workbook has been saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a 2‑D bar chart, enable data labels, and apply a currency number format with two decimal places ("$#,##0.00" and built‑in format index 2) to the labels using Aspose.Cells for C#. The workbook is saved as an Excel file.
    public class BarChartDataLabelCurrencyFormat
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
                worksheet.Cells["B2"].PutValue(1000);
                worksheet.Cells["B3"].PutValue(2000);
                worksheet.Cells["B4"].PutValue(3000);

                // Add a 2‑D bar chart
                int chartIndex = worksheet.Charts.Add(ChartType.Bar, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series
                Series series = chart.NSeries[0];

                // Enable data labels to show the values
                series.DataLabels.ShowValue = true;

                // Set the data label number format to currency with two decimal places
                series.DataLabels.NumberFormat = "$#,##0.00";

                // Optionally, also set the built‑in number format index for currency (2)
                series.DataLabels.Number = 2;

                // Save the workbook to a file
                string outputPath = "BarChartDataLabelCurrency.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            BarChartDataLabelCurrencyFormat.Run();
        }
    }
}
