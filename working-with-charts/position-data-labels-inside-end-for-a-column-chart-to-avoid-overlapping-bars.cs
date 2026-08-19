// Title: Set column chart data labels to InsideEnd in Aspose.Cells (C#) to prevent overlap
// Description: Creates a workbook, adds category and value data, inserts a column chart, binds the series, enables data labels, and positions them at the InsideEnd of each column so the values appear inside the column tops, eliminating label overlap. Saves the file as an XLSX workbook.
// Keywords: Aspose.Cells | C# | .NET | column chart | data label position | InsideEnd | avoid overlapping labels | LabelPositionType | Excel automation | chart label placement
// Common Searches: Aspose.Cells set data label position InsideEnd | C# column chart label inside end | avoid overlapping data labels Aspose.Cells | how to move chart labels inside column Aspose.Cells | Aspose.Cells chart label options C#
// Developer Intent: Place data labels at the inside end of each column in a chart to keep the visualization clear and avoid label collisions.
// Use Cases: Generate Excel reports where column chart labels are anchored inside the column tops for better readability. | Automate dashboards that display dense column charts without label clutter by using the InsideEnd position. | Create reusable chart templates in .NET applications that automatically apply InsideEnd labeling to all column series.
// AI Prompts: Provide C# code that adds a column chart with Aspose.Cells and sets the data label position to InsideEnd. | Explain the different LabelPositionType values in Aspose.Cells and when to choose InsideEnd for column charts. | Show how to modify an existing Aspose.Cells chart to move its data labels to the inside end of each column.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds category and value data, inserts a column chart, binds the series, enables data labels, and positions them at the InsideEnd of each column so the values appear inside the column tops, eliminating label overlap. Saves the file as an XLSX workbook.
class Program
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Position data labels inside the end of each column to avoid overlapping bars
        series.DataLabels.Position = LabelPositionType.InsideEnd;

        // Save the workbook
        workbook.Save("ColumnChartDataLabelsInsideEnd.xlsx");
    }
}
