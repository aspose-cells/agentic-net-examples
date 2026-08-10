// Title: Use a worksheet range as custom data labels in an Aspose.Cells .NET chart
// Description: Creates a workbook, adds categories and values, stores custom label strings in a separate column, builds a column chart, links the series data labels to the range C2:C4 via the LinkedSource property, hides numeric values, sets label font color and position, and saves the file as ChartWithCustomDataLabels.xlsx.
// Keywords: Aspose.Cells | C# | .NET | chart data labels | custom data labels | LinkedSource property | cell range labels | column chart | label formatting | hide numeric values | label font color | label position
// Common Searches: Aspose.Cells custom data labels from cell range | set chart data label source to worksheet cells .NET | display text instead of values on Aspose.Cells chart | change data label font color and position Aspose.Cells | link series data labels to a range in Aspose.Cells
// Developer Intent: Show custom text stored in worksheet cells as the data labels of a chart created with Aspose.Cells for .NET.
// Use Cases: Replace numeric values with product names on a sales column chart. | Show month abbreviations from a separate column as labels on a performance chart. | Apply a specific font color and inside‑end position to custom labels sourced from another range.
// AI Prompts: Generate C# code that binds a cell range to data labels of a pie chart using Aspose.Cells and hides the default values. | Provide an example that sets individual font styles for each custom label based on a second cell range. | Explain how to update the LinkedSource range automatically when new rows are added to the chart data.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Creates a workbook, adds categories and values, stores custom label strings in a separate column, builds a column chart, links the series data labels to the range C2:C4 via the LinkedSource property, hides numeric values, sets label font color and position, and saves the file as ChartWithCustomDataLabels.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate chart data (categories and values)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);

        // Store custom label strings in a separate column
        sheet.Cells["C1"].PutValue("CustomLabel");
        sheet.Cells["C2"].PutValue("First");
        sheet.Cells["C3"].PutValue("Second");
        sheet.Cells["C4"].PutValue("Third");

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data series and category axis
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure data labels to display the custom strings
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = false;       // Hide the default numeric value
        series.DataLabels.ShowCellRange = true;    // Enable showing cell range as label
        series.DataLabels.LinkedSource = "C2:C4";  // Range containing custom text

        // Optional: adjust label appearance
        series.DataLabels.Font.Color = Color.Blue;
        series.DataLabels.Position = LabelPositionType.InsideEnd;

        // Save the workbook with the chart
        workbook.Save("ChartWithCustomDataLabels.xlsx");
    }
}
