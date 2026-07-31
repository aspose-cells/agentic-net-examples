// Title: Create a mixed‑font rich‑text data label for a chart point with Aspose.Cells for .NET
// Description: Shows how to add a column chart, enable a data label on the first point, set custom text, and use the Characters API to format the prefix in large bold blue font and the numeric value in smaller dark‑red font, then save the workbook as an Excel file.
// Keywords: Aspose.Cells | C# chart data label | rich text label | mixed font size | character formatting | ApplyFont | column chart | custom data label | Excel chart styling
// Common Searches: Aspose.Cells set different font sizes in a chart label | C# format part of a chart data label | rich text data label Aspose.Cells example | how to apply bold blue text to chart label prefix | change color of numeric value in Excel chart label using Aspose
// Developer Intent: Add a custom data label with character‑level font size and color to a specific chart point.
// Use Cases: Display a highlighted prefix (e.g., "Val:") in large blue text while showing the value in smaller red text on a sales chart. | Emphasize key metrics in financial dashboards by mixing fonts within the same chart label. | Generate Excel reports where chart labels combine styled headings and numeric values for clearer visual hierarchy.
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart and sets a rich‑text data label on the first point, using different font sizes and colors for the prefix and the number. | Show how to use the Characters method and ApplyFont to format part of a chart data label in blue bold and another part in dark red. | Provide an example that enables a data label, assigns custom text, and applies mixed‑font styling to a specific chart point in an Excel workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to add a column chart, enable a data label on the first point, set custom text, and use the Characters API to format the prefix in large bold blue font and the numeric value in smaller dark‑red font, then save the workbook as an Excel file.
class RichTextDataLabelDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Alpha");
        sheet.Cells["A3"].PutValue("Beta");
        sheet.Cells["A4"].PutValue("Gamma");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(15);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(45);

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Access the first data point and enable its data label
        ChartPoint point = chart.NSeries[0].Points[0];
        point.DataLabels.ShowValue = true;

        // Set custom text for the data label
        point.DataLabels.Text = "Val:15";

        // Apply rich‑text formatting:
        // 1. Make the prefix "Val:" larger, blue and bold
        point.DataLabels.Characters(0, 4).Font.Size = 14;
        point.DataLabels.Characters(0, 4).Font.Color = Color.Blue;
        point.DataLabels.Characters(0, 4).Font.IsBold = true;

        // 2. Make the numeric part smaller and dark red
        int numericStart = 4;
        int numericLength = point.DataLabels.Text.Length - numericStart;
        point.DataLabels.Characters(numericStart, numericLength).Font.Size = 10;
        point.DataLabels.Characters(numericStart, numericLength).Font.Color = Color.DarkRed;

        // Apply the font settings to all child nodes of the label
        point.DataLabels.ApplyFont();

        // Save the workbook
        workbook.Save("RichTextDataLabel.xlsx");
    }
}
