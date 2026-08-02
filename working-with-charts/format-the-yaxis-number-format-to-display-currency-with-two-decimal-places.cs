// Title: Format Chart Y‑Axis as Currency with Two Decimals in Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, inserts a column chart, and applies the number format "$#,##0.00" to the chart's ValueAxis tick labels so the Y‑axis shows monetary values with a dollar sign and two decimal places, then saves the file as YAxisCurrency.xlsx.
// Keywords: Aspose.Cells | C# chart axis formatting | ValueAxis number format | currency format Excel chart | .NET chart tick labels | Y‑axis currency Aspose.Cells | Excel number format C# | chart axis display dollars
// Common Searches: Aspose.Cells set Y axis to currency C# | format chart axis number format .NET | how to display dollars on Excel chart axis using Aspose.Cells | C# change ValueAxis tick labels to $#,##0.00 | apply two‑decimal currency format to chart Y‑axis
// Developer Intent: Apply a currency number format with two decimal places to a chart's Y‑axis in an Aspose.Cells workbook using C#.
// Use Cases: Financial dashboards where chart axes must reflect monetary values. | Sales reports that require clear dollar‑based scaling on column charts. | Budgeting spreadsheets exported to Excel with properly formatted chart axes.
// AI Prompts: Show how to format the Y‑axis of an Aspose.Cells chart as Euro with three decimal places in C#. | Provide C# code to set custom number formats for both X‑axis and Y‑axis tick labels in an Aspose.Cells chart. | Explain using culture‑specific currency symbols for chart axes with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, inserts a column chart, and applies the number format "$#,##0.00" to the chart's ValueAxis tick labels so the Y‑axis shows monetary values with a dollar sign and two decimal places, then saves the file as YAxisCurrency.xlsx.
class FormatYAxisCurrency
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(1234.56);
        sheet.Cells["B3"].PutValue(2345.67);
        sheet.Cells["B4"].PutValue(3456.78);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set Y‑axis (ValueAxis) tick labels to currency format with two decimal places
        chart.ValueAxis.TickLabels.NumberFormat = "$#,##0.00";

        // Save the workbook
        workbook.Save("YAxisCurrency.xlsx");
    }
}
