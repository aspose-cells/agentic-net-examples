// Title: Right‑justify data label text in a horizontal bar chart – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample categories and values, inserts a horizontal bar chart, enables data labels for the first series, and sets the DataLabels.TextHorizontalAlignment property to Right, producing right‑aligned label text before saving the file as an Excel workbook.
// Keywords: Aspose.Cells C# data label alignment | horizontal bar chart label right alignment | TextHorizontalAlignment Right Aspose.Cells | chart data labels alignment .NET | Aspose.Cells chart formatting
// Common Searches: How to right‑align data labels in a horizontal bar chart using Aspose.Cells C# | Aspose.Cells set TextHorizontalAlignment for chart data labels | C# Aspose.Cells change data label alignment bar chart | Right‑justify chart data labels Aspose.Cells .NET | Align data label text to right in Excel chart programmatically
// Developer Intent: Set data label text alignment to right for a horizontal bar chart in Aspose.Cells (C#).
// Use Cases: Generate Excel sales dashboards where bar chart labels need right‑justified values. | Create financial reports with consistent label alignment across multiple bar charts. | Automate chart styling in bulk Excel files to match corporate branding guidelines. | Prepare presentation‑ready Excel files with right‑aligned data labels for readability.
// AI Prompts: Provide C# code using Aspose.Cells to create a horizontal bar chart and set DataLabels.TextHorizontalAlignment to Right. | Show how to apply right‑justified data label alignment to all series in an Aspose.Cells bar chart. | Explain the steps to enable and right‑align data labels in a horizontal bar chart with Aspose.Cells for .NET. | Give a concise example of adjusting chart data label alignment in Aspose.Cells C#.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample categories and values, inserts a horizontal bar chart, enables data labels for the first series, and sets the DataLabels.TextHorizontalAlignment property to Right, producing right‑aligned label text before saving the file as an Excel workbook.
class HorizontalBarChartDataLabelAlignment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a horizontal bar chart
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;

        // Right‑justify the text inside the data labels
        dataLabels.TextHorizontalAlignment = TextAlignmentType.Right;

        // Save the workbook
        workbook.Save("HorizontalBarDataLabelRightAlign.xlsx");
    }
}
