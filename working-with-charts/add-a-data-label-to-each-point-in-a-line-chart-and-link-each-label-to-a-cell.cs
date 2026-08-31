// Title: Add data labels from worksheet cells to each point of a line chart using Aspose.Cells for C#
// AI Prompts: Create a line chart, bind its series to a numeric range, and enable data labels that pull text from a separate cell range. | Hide the default numeric label values and set the label position to appear above each data point in the series. | Link the data labels to a cell range, synchronize number formatting, recalculate the chart, and save the workbook as an XLSX file.
// Common Searches: Aspose.Cells C# line chart data labels linked to cells example | how to display custom text labels on line chart points using Aspose.Cells | remove numeric values from chart series and show cell values as labels in Aspose.Cells | set data label position above points in Aspose.Cells line chart C#
// Tags: Aspose.Cells line chart data labels linked source | C# chart series hide default label values | Aspose.Cells set data label position above points | export chart with custom cell labels to XLSX | Aspose.Cells linked data labels from worksheet cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills columns A‑C with categories, numeric values, and custom label texts, adds a line chart, binds the series to the values and categories, enables data labels, hides the default numeric values, positions the labels above each point, links the labels to cells C2:C4, synchronizes number formatting, forces chart recalculation, and saves the file as LineChart_WithLinkedDataLabels.xlsx.
class AddDataLabelsLinkedToCells
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Column A: Categories (X‑axis)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        // Column B: Values for the line series
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Column C: Text that will be shown as data labels (linked to each point)
        sheet.Cells["C1"].PutValue("Label");
        sheet.Cells["C2"].PutValue("Q1");
        sheet.Cells["C3"].PutValue("Q2");
        sheet.Cells["C4"].PutValue("Q3");

        // Add a line chart
        int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and the category (X‑axis) data
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = false;               // Hide the default numeric value
        series.DataLabels.Position = LabelPositionType.Above; // Position suitable for line charts

        // Link each data label to the corresponding cell in column C
        series.DataLabels.LinkedSource = "C2:C4";
        series.DataLabels.NumberFormatLinked = true;       // Keep number format in sync with source cells

        // Optional: ensure the chart recalculates so that linked labels are applied
        chart.Calculate();

        // Save the workbook
        workbook.Save("LineChart_WithLinkedDataLabels.xlsx");
    }
}
