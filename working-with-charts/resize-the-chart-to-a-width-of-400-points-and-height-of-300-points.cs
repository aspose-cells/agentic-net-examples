// Title: C# – Resize an Aspose.Cells chart to 400 × 300 points
// Description: Creates a workbook, adds sample data, inserts a column chart, then sets ChartObject.Width = 400 and ChartObject.Height = 300 (points) before saving as ResizedChart.xlsx.
// Keywords: Aspose.Cells chart size | ChartObject.Width | ChartObject.Height | C# Excel chart resize | set chart dimensions points
// Common Searches: Aspose.Cells set chart width and height .NET | Resize Excel chart to 400x300 points C# | Change chart object size programmatically Aspose | How to adjust chart dimensions in Aspose.Cells
// Developer Intent: Assign 400 points to the chart's width and 300 points to its height using Aspose.Cells for .NET.
// Use Cases: Generate reports with uniformly sized charts for a clean layout. | Prepare charts for embedding in PDFs or presentations where exact dimensions are required. | Standardize visual elements across multiple workbooks created by an automated process.
// AI Prompts: Write C# code with Aspose.Cells that creates a pie chart and sets its size to 500 × 400 points. | Show how to load an existing workbook, locate a chart, and modify its width and height properties. | Explain the impact of ChartObject.Width and ChartObject.Height on the rendered Excel chart and how points translate to screen pixels.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, then sets ChartObject.Width = 400 and ChartObject.Height = 300 (points) before saving as ResizedChart.xlsx.
class ResizeChartExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Optional: add some data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Resize the chart: width = 400 points, height = 300 points
        chart.ChartObject.Width = 400;   // Width in points
        chart.ChartObject.Height = 300;  // Height in points

        // Save the workbook (lifecycle: save)
        workbook.Save("ResizedChart.xlsx");
    }
}
