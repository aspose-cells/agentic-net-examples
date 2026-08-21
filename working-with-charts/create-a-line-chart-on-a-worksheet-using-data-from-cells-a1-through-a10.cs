// Title: Create a line chart from A1:A10 with Aspose.Cells for .NET (C#)
// Description: This C# example uses Aspose.Cells to generate a new workbook, populate cells A1‑A10 with numeric values, insert a Line chart referencing that range, set an optional title, and save the file as an XLSX workbook.
// Keywords: Aspose.Cells | C# line chart | A1:A10 chart | ChartType.Line | add chart to worksheet | Excel automation | generate line chart programmatically | Aspose.Cells example | save workbook as XLSX
// Common Searches: Aspose.Cells create line chart C# | C# add line chart from range A1:A10 | How to set chart data series Aspose.Cells | Aspose.Cells line chart example .NET | Generate Excel line chart without Excel UI
// Developer Intent: Create a line chart on a worksheet using the values stored in cells A1 through A10.
// Use Cases: Display time‑series data such as sales trends directly from code | Automate chart generation for batch reporting pipelines | Produce consistent visualizations across multiple workbooks in a .NET application | Integrate chart creation into server‑side services that generate Excel files on demand
// AI Prompts: Show how to add multiple data series to the same line chart using Aspose.Cells. | Demonstrate customizing axis labels, line colors, and markers for a line chart in C#. | Provide code to export the created line chart as a PNG image while keeping the workbook unchanged. | Explain how to position the chart dynamically based on worksheet dimensions.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example uses Aspose.Cells to generate a new workbook, populate cells A1‑A10 with numeric values, insert a Line chart referencing that range, set an optional title, and save the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill cells A1 through A10 with sample numeric data
        for (int i = 0; i < 10; i++)
        {
            // Row index i, column index 0 corresponds to column A
            sheet.Cells[i, 0].PutValue(i + 1);
        }

        // Add a line chart to the worksheet (positioned from row 5, column 0 to row 20, column 8)
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart series (A1:A10)
        chart.NSeries.Add("=Sheet1!$A$1:$A$10", true);

        // Optional: give the chart a title
        chart.Title.Text = "Line Chart from A1:A10";

        // Save the workbook to a file
        workbook.Save("LineChart.xlsx", SaveFormat.Xlsx);
    }
}
