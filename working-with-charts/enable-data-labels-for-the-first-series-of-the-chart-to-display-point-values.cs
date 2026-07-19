// Title: Enable Data Labels (Show Values) for the First Series in an Aspose.Cells Column Chart (C#)
// Description: Creates a workbook, inserts sample data, adds a column chart, defines the series range, and activates DataLabels.ShowValue for the first series. The chart is recalculated and saved, resulting in each column displaying its numeric value.
// Keywords: Aspose.Cells C# chart data labels | ShowValue property Aspose.Cells | enable data labels column chart | display point values Aspose.Cells | .NET Excel chart series labels | Aspose.Cells chart customization
// Common Searches: Aspose.Cells C# show values on chart series | how to enable data labels for first series Aspose.Cells | chart.ShowValue Aspose.Cells example | display data labels in column chart using Aspose.Cells
// Developer Intent: Turn on data labels for the first series of a column chart so each column shows its underlying value.
// Use Cases: Add sales figures directly on a column chart in an automated Excel report. | Highlight KPI numbers in a performance dashboard generated with Aspose.Cells. | Improve readability of exported Excel charts for end‑users by showing point values.
// AI Prompts: Generate C# code with Aspose.Cells that enables data labels for every series in a line chart and customizes font size and color. | Show how to set ShowValue = true for the second series of a pie chart and then recalculate the chart. | Explain the steps to refresh chart data labels after updating the source data in an Aspose.Cells workbook.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, inserts sample data, adds a column chart, defines the series range, and activates DataLabels.ShowValue for the first series. The chart is recalculated and saved, resulting in each column displaying its numeric value.
class EnableDataLabels
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series to display point values
        Series firstSeries = chart.NSeries[0];
        firstSeries.DataLabels.ShowValue = true;

        // Recalculate the chart (optional but ensures proper layout)
        chart.Calculate();

        // Save the workbook
        workbook.Save("EnableDataLabels.xlsx");
    }
}
