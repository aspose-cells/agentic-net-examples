// Title: Aspose.Cells for .NET – Enable ‘Label Contains – Value From Cells’ on a Pie Chart and Link Labels to Worksheet Cells
// Description: C# code that creates a workbook, adds category, value, and custom label data, inserts a pie chart, defines the series range, activates the ShowCellRange option, links each slice label to cells C2:C4 via the LinkedSource property, optionally keeps the number format linked, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# pie chart | ShowCellRange | DataLabels.LinkedSource | custom chart labels | link chart labels to cells | pie chart data labels from cells | Aspose.Cells .NET | Excel chart label binding | Series.DataLabels
// Common Searches: Aspose.Cells link pie chart labels to cells | ShowCellRange option C# Aspose | How to bind chart data labels to a cell range in Aspose.Cells | Custom labels for pie chart using Aspose.Cells .NET | Series.DataLabels.ShowCellRange example
// Developer Intent: Activate ShowCellRange for a pie‑chart series and bind each label to a specific worksheet cell range.
// Use Cases: Generate Excel reports where pie‑chart slice labels display dynamic text stored in worksheet cells. | Create dashboards that automatically update chart labels when underlying cell values or formulas change. | Preserve the original number formatting of label cells when rendering the chart in Excel.
// AI Prompts: Write C# code with Aspose.Cells that enables ShowCellRange for a chart series and links the data labels to a given cell range. | Show an example of linking pie‑chart labels to cells while keeping the number format linked using Aspose.Cells for .NET. | Explain how to modify the cells containing custom labels and refresh the chart so the changes appear in an existing workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# code that creates a workbook, adds category, value, and custom label data, inserts a pie chart, defines the series range, activates the ShowCellRange option, links each slice label to cells C2:C4 via the LinkedSource property, optionally keeps the number format linked, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["A4"].PutValue("Cherry");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(45);
        sheet.Cells["B4"].PutValue(25);

        // Cells that contain the custom labels to be linked
        sheet.Cells["C1"].PutValue("Label");
        sheet.Cells["C2"].PutValue("Apple - 30 units");
        sheet.Cells["C3"].PutValue("Banana - 45 units");
        sheet.Cells["C4"].PutValue("Cherry - 25 units");

        // Add a pie chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable "Label Contains – Value From Cells" and link labels to the cells in column C
        Series series = chart.NSeries[0];
        series.DataLabels.ShowCellRange = true;          // activates the option
        series.DataLabels.LinkedSource = "C2:C4";        // links each label to its corresponding cell
        series.DataLabels.NumberFormatLinked = true;    // optional: keep number format linked

        // Save the workbook
        workbook.Save("PieChartLabelFromCells.xlsx");
    }
}
