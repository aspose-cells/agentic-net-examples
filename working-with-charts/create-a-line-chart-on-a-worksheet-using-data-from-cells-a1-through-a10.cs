// Title: Add a Line Chart from A1:A10 with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to generate a new workbook, populate cells A1‑A10 with numeric values, insert a Line chart, bind the chart to the A1:A10 range, set a title, position the chart on the sheet, and save the file as XLSX using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# line chart | ChartType.Line | SetChartDataRange | Excel chart from range | A1:A10 chart data | Create workbook Aspose | Add chart to worksheet | Save XLSX Aspose.Cells
// Common Searches: Aspose.Cells line chart example C# | How to bind chart to A1:A10 in Aspose.Cells | Create line chart in Excel using Aspose.Cells .NET | C# add chart to worksheet Aspose | Set chart data range vertically Aspose.Cells
// Developer Intent: Generate and persist a line chart that visualizes the values in cells A1‑A10 of a .NET workbook.
// Use Cases: Plot monthly sales figures stored in a single column to show trends. | Display daily temperature readings from a sensor array for quick analysis. | Automate a financial report that inserts a line chart for a predefined data block and exports it as XLSX.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart for the range B2:B15, include a custom title and legend. | Show how to update an existing Aspose.Cells line chart to use a new data range and change its type to a scatter chart. | Provide a snippet that positions a line chart at a specific cell range and applies a predefined chart style.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to generate a new workbook, populate cells A1‑A10 with numeric values, insert a Line chart, bind the chart to the A1:A10 range, set a title, position the chart on the sheet, and save the file as XLSX using Aspose.Cells for C#.
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

        // Add a line chart to the worksheet.
        // Parameters: chart type, top row, left column, bottom row, right column
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 1, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (A1:A10) and plot vertically
        chart.SetChartDataRange("A1:A10", true);

        // Optional: set a title for the chart
        chart.Title.Text = "Line Chart from A1:A10";

        // Save the workbook to a file
        workbook.Save("LineChartFromA1toA10.xlsx", SaveFormat.Xlsx);
    }
}
