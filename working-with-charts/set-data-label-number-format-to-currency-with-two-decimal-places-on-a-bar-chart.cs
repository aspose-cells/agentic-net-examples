// Title: Aspose.Cells for .NET – Set Bar Chart Data Labels to Currency ($#,##0.00) with Two Decimals
// Description: Creates a workbook, adds sample categories and values, inserts a bar chart, enables data labels, applies the number format "$#,##0.00" to show currency with two decimal places, and saves the file as BarChartDataLabelsCurrency.xlsx using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | Bar chart | Data labels | Currency format | NumberFormat | Excel chart formatting | Financial chart | Aspose.Cells example | Chart series label formatting
// Common Searches: Aspose.Cells set chart data label currency format C# | How to format bar chart labels as $ with two decimals in .NET | Change number format of chart series data labels Aspose.Cells | C# Aspose.Cells bar chart currency data labels | Apply $#,##0.00 format to Excel chart labels using Aspose
// Developer Intent: Apply a currency number format with two decimal places to the data labels of a bar chart in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Financial dashboards where bar chart columns display sales amounts like $1,234.56. | Budget reports that need monetary values shown directly on chart data labels. | Automated generation of cost‑analysis workbooks with standardized currency formatting on charts.
// AI Prompts: Generate C# code with Aspose.Cells to set bar chart data label number format to "$#,##0.00". | Explain how to customize chart series data label formatting in Aspose.Cells, including currency examples. | Show step‑by‑step how to apply a custom currency format to Excel chart labels and save the workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample categories and values, inserts a bar chart, enables data labels, applies the number format "$#,##0.00" to show currency with two decimal places, and saves the file as BarChartDataLabelsCurrency.xlsx using Aspose.Cells in C#.
class SetDataLabelCurrency
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the bar chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["B3"].PutValue(2000);
        sheet.Cells["B4"].PutValue(3000);

        // Add a bar chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Enable data labels and set their number format to currency with two decimals
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;
        series.DataLabels.NumberFormat = "$#,##0.00";

        // Save the workbook to a file
        workbook.Save("BarChartDataLabelsCurrency.xlsx");
    }
}
