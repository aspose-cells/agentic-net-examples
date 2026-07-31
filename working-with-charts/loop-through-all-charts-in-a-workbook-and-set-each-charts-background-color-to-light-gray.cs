// Title: C# – Set All Chart Backgrounds to Light Gray with Aspose.Cells
// Description: Creates a workbook, adds sample data and a chart, then iterates through every worksheet and chart to apply a LightGray background to both the ChartArea and PlotArea before saving the file.
// Keywords: Aspose.Cells chart background | C# set chart area color | loop through charts workbook | light gray Excel chart | Aspose.Cells ChartArea PlotArea | .NET Excel chart styling
// Common Searches: how to change chart background color in Aspose.Cells C# | set all chart areas to light gray programmatically | loop through worksheets and modify chart colors Aspose | apply uniform chart background across Excel file | Aspose.Cells change plot area color for multiple charts
// Developer Intent: Apply a LightGray background to the ChartArea and PlotArea of every chart in a workbook using Aspose.Cells for .NET.
// Use Cases: Standardize chart appearance across a multi‑sheet report before distribution. | Create a corporate template where all charts share a neutral light‑gray background for consistent printing. | Batch‑update existing Excel workbooks to align chart colors with branding guidelines.
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheets and sets each chart's ChartArea and PlotArea background to a specified color. | Show an example that includes error handling when modifying chart backgrounds for different chart types in a workbook. | Explain how to retrieve and change the background color of charts in Aspose.Cells without affecting other chart properties.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data and a chart, then iterates through every worksheet and chart to apply a LightGray background to both the ChartArea and PlotArea before saving the file.
class SetChartBackground
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a sample chart (optional, just to demonstrate)
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Loop through all worksheets and their charts
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                // Set the background color of the chart area to light gray
                ch.ChartArea.Area.BackgroundColor = Color.LightGray;

                // Also set the plot area background to light gray (optional)
                ch.PlotArea.Area.BackgroundColor = Color.LightGray;
            }
        }

        // Save the workbook
        workbook.Save("AllChartsBackgroundLightGray.xlsx");
    }
}
