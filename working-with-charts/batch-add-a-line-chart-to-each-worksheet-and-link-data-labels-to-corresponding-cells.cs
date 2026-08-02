// Title: Add a Line Chart with Linked Data Labels to Every Worksheet using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data to each sheet, inserts a LineWithDataMarkers chart, assigns series values and categories, links the chart's data‑label text to a separate label column, syncs number formatting, and saves the file as BatchLineCharts.xlsx.
// Keywords: Aspose.Cells line chart C# | batch chart creation Aspose.Cells | linked data labels Aspose.Cells | multiple worksheets chart .NET | ChartType.LineWithDataMarkers | NSeries data source | DataLabels.LinkedSource | programmatic Excel chart | dynamic chart labels
// Common Searches: how to add a line chart to each sheet with Aspose.Cells | link chart data labels to cells in C# Aspose.Cells | batch generate charts across worksheets .NET | sync number format for chart data labels Aspose.Cells | Aspose.Cells example for LineWithDataMarkers
// Developer Intent: Programmatically insert a line chart on every worksheet and bind its data‑label text to corresponding cells so the labels update automatically with source data.
// Use Cases: Create regional sales‑trend line charts in a multi‑sheet workbook, linking each point’s label to a description column that changes with the data. | Automate a monthly performance report where each sheet receives a line chart whose labels are tied to comment cells, ensuring real‑time updates when values are edited. | Build a template workbook that pre‑populates line charts with linked labels, allowing end users to add rows and have both the chart and its labels expand without additional coding.
// AI Prompts: Generate C# code with Aspose.Cells that adds a bar chart to every worksheet and links its data labels to a separate text column. | Show how to change the chart type to Area and customize the data‑label font, color, and background using Aspose.Cells. | Explain how to adjust the linked source range for data labels after inserting new rows into the worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data to each sheet, inserts a LineWithDataMarkers chart, assigns series values and categories, links the chart's data‑label text to a separate label column, syncs number formatting, and saves the file as BatchLineCharts.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Ensure there are multiple worksheets for the batch operation
        // Add a second worksheet for demonstration
        workbook.Worksheets.Add();

        // Iterate through each worksheet and add a line chart with linked data labels
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // ----- Sample data setup (Category, Value, Label) -----
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Label");

            for (int i = 1; i <= 5; i++)
            {
                sheet.Cells[$"A{i + 1}"].PutValue($"Item {i}");
                sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
                sheet.Cells[$"C{i + 1}"].PutValue($"Lbl {i}");
            }

            // ----- Add a line chart to the current worksheet -----
            // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
            int chartIndex = sheet.Charts.Add(ChartType.LineWithDataMarkers, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Build range strings that include the worksheet name
            string sheetName = sheet.Name;
            string valuesRange   = $"='{sheetName}'!$B$2:$B$6";
            string categoryRange = $"='{sheetName}'!$A$2:$A$6";
            string labelRange    = $"='{sheetName}'!$C$2:$C$6";

            // Set the data source for the chart
            chart.NSeries.Add(valuesRange, true);          // Y‑values
            chart.NSeries.CategoryData = categoryRange;    // X‑axis categories

            // ----- Configure data labels to link to cells -----
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;            // Show the value
            series.DataLabels.LinkedSource = labelRange;   // Link label text to cells
            series.DataLabels.NumberFormatLinked = true;   // Keep number format in sync
        }

        // Save the workbook with all charts added
        workbook.Save("BatchLineCharts.xlsx", SaveFormat.Xlsx);
    }
}
