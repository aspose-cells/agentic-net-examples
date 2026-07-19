// Title: Right‑Align Data Labels in a Horizontal Bar Chart with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a horizontal bar chart, enable data labels, and set the TextHorizontalAlignment property to Right so the label text appears flush with the bar edge. The example saves the result as an Excel file.
// Keywords: Aspose.Cells chart data label alignment | C# right align data labels | horizontal bar chart label formatting | TextHorizontalAlignment Right | .NET Excel chart customization | Aspose.Cells label positioning
// Common Searches: Aspose.Cells set data label alignment right | C# horizontal bar chart right‑aligned labels | How to right‑justify chart data labels in .NET | TextHorizontalAlignment property Aspose.Cells example | Align bar chart data labels to the right
// Developer Intent: Apply right‑justified alignment to the text of data labels on a horizontal bar chart using Aspose.Cells for .NET.
// Use Cases: Financial dashboards where bar values need to line up on the right side for quick comparison. | Sales performance reports that require right‑aligned labels to match corporate styling guidelines. | Automated Excel exports that include bar charts with labels positioned at the bar end for improved readability.
// AI Prompts: Generate C# code with Aspose.Cells that right‑aligns data labels on a horizontal bar chart. | Explain the effect of the TextHorizontalAlignment property on chart data labels and list all available alignment options in Aspose.Cells. | Show how to apply right‑aligned data labels to multiple series in a bar chart using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a horizontal bar chart, enable data labels, and set the TextHorizontalAlignment property to Right so the label text appears flush with the bar edge. The example saves the result as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a horizontal bar chart
        int chartIndex = worksheet.Charts.Add(ChartType.Bar, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        chart.NSeries[0].DataLabels.ShowValue = true;

        // Right‑justify the data label text
        chart.NSeries[0].DataLabels.TextHorizontalAlignment = TextAlignmentType.Right;

        // Save the workbook
        workbook.Save("HorizontalBarChartDataLabelRightAlign.xlsx");
    }
}
