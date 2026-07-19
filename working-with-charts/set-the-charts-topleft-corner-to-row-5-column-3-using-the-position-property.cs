// Title: Aspose.Cells C# – Set Chart Top‑Left Corner to Row 5, Column 3 Using Chart.Move
// Description: Shows how to create a workbook, add a column chart, and reposition its top‑left corner to row 5, column 3 while keeping the original size (bottom‑right at row 20, column 5) via the Chart.Move method, then save as ChartPositionDemo.xlsx.
// Keywords: Aspose.Cells | C# | Chart.Move | chart position | set chart top left cell | move chart row column | Aspose.Cells chart placement | Excel chart reposition | programmatic chart location | Aspose.Cells API
// Common Searches: Aspose.Cells move chart to specific cell | C# set chart top left corner row 5 column 3 | How to change chart location in Aspose.Cells | Chart.Move parameters explanation | Reposition Excel chart with Aspose.Cells
// Developer Intent: Move an existing chart so its top‑left corner aligns with a given row and column while preserving its dimensions.
// Use Cases: Align a chart with a header row after inserting new rows | Create a dashboard where each chart starts at a predefined cell | Adjust chart placement based on user‑selected cell coordinates without altering size
// AI Prompts: Write C# code using Aspose.Cells to move a chart to row 8, column 2 while keeping its current width and height. | Explain the four parameters of Chart.Move and how they map to top‑left and bottom‑right cell indices. | Provide a method that calculates the bottom‑right cell indices automatically after changing a chart’s top‑left position.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add a column chart, and reposition its top‑left corner to row 5, column 3 while keeping the original size (bottom‑right at row 20, column 5) via the Chart.Move method, then save as ChartPositionDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Fruits");
        worksheet.Cells["A3"].PutValue("Vegetables");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(50);
        worksheet.Cells["B3"].PutValue(30);

        // Add a column chart; initial position is rows 10‑20, columns 0‑5
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 10, 0, 20, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Move the chart so that its top‑left corner is at row 5, column 3
        // Keep the original bottom‑right corner (rows 20, column 5) for size consistency
        chart.Move(5, 3, 20, 5);

        // Save the workbook
        workbook.Save("ChartPositionDemo.xlsx");
    }
}
