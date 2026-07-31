// Title: C# – Add Linked Data Labels to Every Point of a Line Chart with Aspose.Cells
// Description: Creates a workbook, fills category and value ranges, inserts a line chart, enables data labels, then links each point's label to a specific cell (e.g., C2‑C4). The labels inherit the cell's number format, are positioned above the points, and the workbook is saved as an Excel file.
// Keywords: Aspose.Cells line chart data labels C# | link chart label to cell Aspose.Cells | add point labels line chart .NET | chart data label position Aspose | Excel line chart label linking
// Common Searches: Aspose.Cells link data label to worksheet cell | C# add individual labels to line chart points | set label position above line chart points Aspose | inherit number format for chart data labels Aspose.Cells | save line chart with linked labels in .NET
// Developer Intent: Add a data label to each point of a line chart and bind each label to a corresponding worksheet cell.
// Use Cases: Display custom target names (e.g., Q1 Target) beside each data point in a sales trend line chart. | Keep chart labels automatically updated when the source cells change. | Improve readability by positioning labels above points in multi‑series line charts.
// AI Prompts: Write C# code using Aspose.Cells to add a data label to every point of a line chart and link each label to cells in column C. | Explain how to change the position of chart point labels and enable number‑format linking in Aspose.Cells. | Show how to update linked label cells dynamically after modifying the chart data series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills category and value ranges, inserts a line chart, enables data labels, then links each point's label to a specific cell (e.g., C2‑C4). The labels inherit the cell's number format, are positioned above the points, and the workbook is saved as an Excel file.
class AddDataLabelsToLineChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Categories
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        // Values for the line series
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Cells that will be linked to data labels
        sheet.Cells["C2"].PutValue("Q1 Target");
        sheet.Cells["C3"].PutValue("Q2 Target");
        sheet.Cells["C4"].PutValue("Q3 Target");

        // Add a line chart
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the series (required to access point labels)
        chart.NSeries[0].DataLabels.ShowValue = true;

        // Iterate through each point and link its label to the corresponding cell in column C
        Series series = chart.NSeries[0];
        for (int i = 0; i < series.Points.Count; i++)
        {
            ChartPoint point = series.Points[i];
            // Link the data label to the cell (e.g., C2, C3, C4)
            string cellAddress = $"C{2 + i}";
            point.DataLabels.LinkedSource = cellAddress;
            point.DataLabels.NumberFormatLinked = true; // keep number format in sync if needed
            point.DataLabels.Position = LabelPositionType.Above; // position suitable for line chart
        }

        // Save the workbook
        workbook.Save("LineChartWithLinkedDataLabels.xlsx");
    }
}
