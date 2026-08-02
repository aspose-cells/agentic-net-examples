// Title: Use a cell range as custom data labels for the first series of a column chart – Aspose.Cells C# example
// Description: Creates a workbook, fills columns A‑C with categories, values and label text, adds a column chart, binds series values (B2:B4) and categories (A2:A4), enables data labels, sets ShowCellRange = true, links labels to C2:C4, optionally shows the numeric value, styles the font, and saves the file as ChartWithCellRangeDataLabels.xlsx.
// Keywords: Aspose.Cells | C# chart data labels | ShowCellRange | LinkedSource | custom chart labels | cell range labels | column chart series | Aspose.Cells .NET | Excel chart labeling | programmatic chart labels
// Common Searches: Aspose.Cells set data labels from cells | C# link chart series labels to worksheet range | How to use ShowCellRange in Aspose.Cells | Display custom text on column chart data labels .NET | Aspose.Cells chart label LinkedSource example | Create chart with cell‑based data labels using Aspose.Cells
// Developer Intent: The developer wants to display custom label text taken from a worksheet range for the first series of a column chart.
// Use Cases: Sales dashboard where each column shows a formatted label like "100 units" stored in a separate column. | Financial report that automatically updates chart labels when the source cells are edited. | Presentation slide with a column chart whose data labels are editable directly in the worksheet by non‑technical users.
// AI Prompts: Write C# code with Aspose.Cells to link data labels of a line chart series to a specified cell range and customize the label font color. | Show how to hide numeric values and display only custom text from cells for multiple series in a bar chart using Aspose.Cells. | Explain how to change the LinkedSource range for data labels after the workbook has been created and saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Creates a workbook, fills columns A‑C with categories, values and label text, adds a column chart, binds series values (B2:B4) and categories (A2:A4), enables data labels, sets ShowCellRange = true, links labels to C2:C4, optionally shows the numeric value, styles the font, and saves the file as ChartWithCellRangeDataLabels.xlsx.
class EnableDataLabelsFromCellRange
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);

        // Custom label texts stored in another column
        sheet.Cells["C1"].PutValue("Label");
        sheet.Cells["C2"].PutValue("100 units");
        sheet.Cells["C3"].PutValue("200 units");
        sheet.Cells["C4"].PutValue("300 units");

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the first series (values)
        chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
        // Set category (X) axis data
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

        // Enable data labels for the first series
        Series firstSeries = chart.NSeries[0];
        firstSeries.DataLabels.ShowCellRange = true;          // Use cell range as label source
        firstSeries.DataLabels.LinkedSource = "C2:C4";        // Reference to custom label cells
        firstSeries.DataLabels.ShowValue = true;              // Optional: also show the numeric value
        firstSeries.DataLabels.Font.Color = Color.Blue;       // Example styling

        // Save the workbook (save rule)
        workbook.Save("ChartWithCellRangeDataLabels.xlsx", SaveFormat.Xlsx);
    }
}
