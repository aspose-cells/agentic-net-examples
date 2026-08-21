// Title: Create and embed a column chart from XML‑mapped totals using Aspose.Cells in C#
// Description: Demonstrates how to build a new workbook, populate cells with totals derived from an XML map, add a column chart (rows 5‑20, columns 0‑8), bind it to the range A1:B4, set a title and legend, force layout calculation, and save the file as TotalsChart.xlsx.
// Keywords: Aspose.Cells | C# chart example | XML map to Excel | column chart embedding | .NET Excel chart | set chart data range | calculate chart layout | save workbook as XLSX
// Common Searches: Aspose.Cells create column chart from XML mapped cells | C# embed chart in Excel workbook using Aspose.Cells | How to set chart data range A1:B4 in Aspose.Cells | Calculate chart before saving Aspose.Cells .NET | Generate Excel chart from XML map with Aspose
// Developer Intent: Generate a column chart from totals extracted via an XML map and embed it directly into the worksheet.
// Use Cases: Show sales totals imported from an XML file in a visual chart within an automated report. | Build a dashboard that updates its chart automatically when the underlying XML data changes. | Create a printable Excel summary that combines tabular XML data with a chart for quick insight.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart from XML‑mapped range A1:C10 and embed it in the worksheet. | Show how to bind an Aspose.Cells chart to a dynamic named range that reflects XML map updates. | Explain how to customize colors and data labels for a column chart generated from XML‑derived totals.

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to build a new workbook, populate cells with totals derived from an XML map, add a column chart (rows 5‑20, columns 0‑8), bind it to the range A1:B4, set a title and legend, force layout calculation, and save the file as TotalsChart.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with sample totals that would come from mapped XML nodes
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Total");
        sheet.Cells["A2"].PutValue("Item1");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Item2");
        sheet.Cells["B3"].PutValue(85);
        sheet.Cells["A4"].PutValue("Item3");
        sheet.Cells["B4"].PutValue(150);

        // Add a column chart to the worksheet (positioned from row 5, column 0 to row 20, column 8)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart (vertical series)
        chart.SetChartDataRange("A1:B4", true);

        // Configure basic chart properties
        chart.Title.Text = "Totals from XML Nodes";
        chart.ShowLegend = true;

        // Ensure the chart layout is calculated before saving
        chart.Calculate();

        // Save the workbook with the embedded chart
        workbook.Save("TotalsChart.xlsx", SaveFormat.Xlsx);
    }
}
