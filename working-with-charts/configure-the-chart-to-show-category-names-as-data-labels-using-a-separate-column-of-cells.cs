// Title: Aspose.Cells C# – Show custom text labels from a separate column in a column chart
// Description: Creates a workbook, fills columns with categories, values and label text, adds a column chart, links the series to the value range, sets the X‑axis to the category range, and configures the series to display labels from another cell range while hiding numeric values and applying styling.
// Keywords: Aspose.Cells | C# chart | column chart | custom text labels | linked source for DataLabels | ShowCellRange | Hide chart values | LabelPositionType | DataLabels.Font.Color | Excel chart labeling
// Common Searches: Aspose.Cells set data labels from a cell range | C# chart custom labels Aspose | Link chart labels to another column in Aspose.Cells | Hide numeric values and show text labels in Excel chart | Change data label position Aspose.Cells
// Developer Intent: Replace the default numeric data labels on a column chart with text taken from a separate worksheet column.
// Use Cases: Build a sales dashboard where product names from a description column appear as labels on each bar. | Create a performance report that annotates each data point with a comment stored in an adjacent column. | Export an Excel chart with styled, positioned text labels while suppressing the underlying values for a cleaner visual.
// AI Prompts: Generate C# code using Aspose.Cells to attach a column chart's DataLabels to a specific cell range, hide the values, and set a custom font color. | Explain how to move data labels to a different position and update the linked source range after inserting new rows. | Provide steps to disable default category names and numeric values, showing only custom text from another column in an Aspose.Cells chart.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills columns with categories, values and label text, adds a column chart, links the series to the value range, sets the X‑axis to the category range, and configures the series to display labels from another cell range while hiding numeric values and applying styling.
class ChartCategoryLabelsExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data:
        // Column A – Category names (used for X‑axis)
        // Column B – Values for the series
        // Column C – Custom labels that will be shown as data labels
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Alpha");
        sheet.Cells["A3"].PutValue("Beta");
        sheet.Cells["A4"].PutValue("Gamma");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        sheet.Cells["C1"].PutValue("Label");
        sheet.Cells["C2"].PutValue("First");
        sheet.Cells["C3"].PutValue("Second");
        sheet.Cells["C4"].PutValue("Third");

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and the category (X‑axis) data
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure data labels to use the custom label column (C2:C4)
        Series series = chart.NSeries[0];
        series.DataLabels.ShowCellRange = true;    // Enable using a cell range as label source
        series.DataLabels.LinkedSource = "C2:C4";  // Reference to the custom label cells
        series.DataLabels.ShowValue = false;       // Hide the numeric value
        series.DataLabels.ShowCategoryName = false;
        series.DataLabels.Position = LabelPositionType.InsideEnd; // Optional positioning
        series.DataLabels.Font.Color = Color.DarkBlue;            // Optional styling

        // Save the workbook
        workbook.Save("ChartWithCustomCategoryLabels.xlsx", SaveFormat.Xlsx);
    }
}
