// Title: Add a line chart to an Excel worksheet from cells A1‑A10 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that fills A1‑A10 with numeric values, creates a line chart based on that range, sets a chart title, and saves the workbook as an XLSX file using Aspose.Cells. | Show how to position a line chart on a worksheet and define its data source with Aspose.Cells in C#.
// Common Searches: asp.net create line chart from A1:A10 using Aspose.Cells | c# Aspose.Cells set chart data range A1 to A10 | how to add a line chart to a worksheet with Aspose.Cells .NET | Aspose.Cells line chart positioning rows columns C# | save workbook with chart as XLSX using Aspose.Cells
// Tags: Aspose.Cells line chart creation | Aspose.Cells set chart data range | Aspose.Cells chart positioning worksheet | Aspose.Cells export workbook with chart to XLSX | C# populate cells A1-A10 numeric values

using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, writes numbers 1‑10 into cells A1‑A10, adds a line chart that uses this range as its data source, sets a chart title, positions the chart on the sheet, and saves the file as LineChart.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells A1 through A10 with sample numeric data
        for (int i = 0; i < 10; i++)
        {
            // Row index i, column index 0 corresponds to column A
            sheet.Cells[i, 0].PutValue(i + 1);
        }

        // Add a line chart to the worksheet.
        // Parameters: chart type, top row, left column, bottom row, right column
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (vertical series from A1 to A10)
        chart.SetChartDataRange("A1:A10", true);

        // Optional: set a chart title
        chart.Title.Text = "Line Chart from A1:A10";

        // Save the workbook to a file
        workbook.Save("LineChart.xlsx", SaveFormat.Xlsx);
    }
}
