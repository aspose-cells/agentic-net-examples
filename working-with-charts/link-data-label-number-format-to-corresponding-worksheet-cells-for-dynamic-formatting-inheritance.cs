// Title: Link chart data label number format to worksheet cells in Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add a column chart, and set the series so data labels pull their text from cells C2:C3 and automatically inherit those cells' number format via the NumberFormatLinked property. Includes optional font styling and saves the workbook.
// Keywords: Aspose.Cells | C# chart data labels | NumberFormatLinked | linked source data labels | dynamic label formatting | inherit number format | Excel chart formatting .NET | column chart data labels | currency label formatting | Aspose.Cells chart example
// Common Searches: Aspose.Cells link data label to cells | How to inherit number format for chart labels in Aspose.Cells | NumberFormatLinked property C# example | Chart data labels use linked source Aspose.Cells | Dynamic chart label formatting .NET
// Developer Intent: The developer wants chart data labels to automatically reflect the number format defined in a specific worksheet range, eliminating the need to update formatting code when the cell format changes.
// Use Cases: Financial reports where currency symbols and decimal places are defined in cells and automatically applied to chart labels. | Dashboards that reuse a single formatting range for multiple charts to guarantee consistent appearance. | User‑editable spreadsheets where changing the cell format instantly updates the chart labels. | Localized Excel files that adapt number formats to regional settings via linked cells.
// AI Prompts: Write C# code using Aspose.Cells to link chart data labels to a cell range and enable NumberFormatLinked. | Explain how DataLabels.LinkedSource and NumberFormatLinked work together in Aspose.Cells. | Show how modifying the number format in cells C2:C3 updates chart labels automatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Shows how to create a workbook, add a column chart, and set the series so data labels pull their text from cells C2:C3 and automatically inherit those cells' number format via the NumberFormatLinked property. Includes optional font styling and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(1234.56);
        sheet.Cells["B3"].PutValue(7890.12);
        // Cells C2:C3 contain formatted strings that we want the data labels to inherit
        sheet.Cells["C2"].PutValue("1,234.56 USD");
        sheet.Cells["C3"].PutValue("7,890.12 USD");

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the series and the category axis
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Configure data labels to use the linked source and inherit number format
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;               // Show the numeric value
        series.DataLabels.LinkedSource = "C2:C3";          // Link to cells with formatted text
        series.DataLabels.NumberFormatLinked = true;      // Inherit number format from linked cells

        // Optional: customize label appearance
        series.DataLabels.Font.Color = Color.Blue;
        series.DataLabels.Font.Size = 10;

        // Save the workbook
        workbook.Save("DataLabelsLinkedNumberFormat.xlsx");
    }
}
