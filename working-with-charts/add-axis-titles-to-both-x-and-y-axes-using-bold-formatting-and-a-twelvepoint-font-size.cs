// Title: Set bold 12‑pt X‑ and Y‑axis titles in an Aspose.Cells column chart (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, and formats both the Category (X) and Value (Y) axis titles to be visible, bold, and 12‑point before saving as ChartWithAxisTitles.xlsx.
// Keywords: Aspose.Cells C# chart axis title | set axis label font Aspose.Cells | bold axis title 12pt | column chart axis formatting .NET | Excel chart title styling Aspose
// Common Searches: Aspose.Cells set X axis title bold | How to add Y axis label in C# Aspose.Cells | Change chart axis font size Aspose.Cells .NET | Make chart axis titles visible Aspose | Format chart axis titles programmatically
// Developer Intent: Apply bold, 12‑point titles to both axes of a chart using Aspose.Cells for .NET.
// Use Cases: Standardize axis labeling across corporate Excel reports. | Improve chart readability in automated financial dashboards. | Enforce branding guidelines that require bold 12‑pt axis titles.
// AI Prompts: Generate C# code that sets italic 14‑pt, colored axis titles for all charts in a workbook using Aspose.Cells. | Provide a loop that iterates through every chart in a workbook and applies bold 12‑pt titles to both axes. | Explain how to hide axis titles conditionally based on chart type with Aspose.Cells in C#.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, and formats both the Category (X) and Value (Y) axis titles to be visible, bold, and 12‑point before saving as ChartWithAxisTitles.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure X‑axis (CategoryAxis) title
        chart.CategoryAxis.Title.Text = "Categories";
        chart.CategoryAxis.Title.IsVisible = true;
        chart.CategoryAxis.Title.Font.IsBold = true;
        chart.CategoryAxis.Title.Font.Size = 12;

        // Configure Y‑axis (ValueAxis) title
        chart.ValueAxis.Title.Text = "Values";
        chart.ValueAxis.Title.IsVisible = true;
        chart.ValueAxis.Title.Font.IsBold = true;
        chart.ValueAxis.Title.Font.Size = 12;

        // Save the workbook with the chart
        workbook.Save("ChartWithAxisTitles.xlsx");
    }
}
