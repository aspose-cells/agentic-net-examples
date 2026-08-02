// Title: Aspose.Cells .NET: Create a Tornado (Stacked Bar) Chart with Custom Colors and Export to PNG
// Description: C# example that builds a workbook, fills A1:B5 with region names and sales figures (including negative values), adds a stacked‑bar chart styled as a tornado diagram, assigns a unique solid‑fill color to each bar, recalculates the layout, exports the chart as a PNG image, and saves the workbook as XLSX.
// Keywords: Aspose.Cells | C# | .NET | tornado chart | stacked bar chart | custom bar colors | export chart to PNG | chart image generation | sales data visualization | chart data range | GitHub example
// Common Searches: Aspose.Cells create tornado chart C# | how to set individual bar colors in Aspose.Cells stacked bar chart | export Aspose.Cells chart as PNG image | set chart data range with headers Aspose.Cells | C# example for custom colored chart points Aspose.Cells
// Developer Intent: Generate a tornado‑style stacked bar chart from sales data, color each bar separately, and save the chart as a PNG file (plus the workbook) using Aspose.Cells for .NET.
// Use Cases: Show regional sales gaps in a management report with a side‑by‑side positive/negative bar layout. | Create color‑coded performance bars for a dashboard that instantly distinguishes gains from losses. | Automate production of chart images for newsletters or email alerts directly from workbook data.
// AI Prompts: Write C# code using Aspose.Cells to build a tornado chart from range A1:B5, apply distinct solid‑fill colors to each point, and export the chart as a PNG file. | Explain how to add data labels, adjust chart dimensions, and improve image resolution before exporting the chart image in the provided Aspose.Cells sample. | Provide steps to switch the chart to a horizontal bar type and dynamically set bar colors based on whether sales values are positive or negative.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that builds a workbook, fills A1:B5 with region names and sales figures (including negative values), adds a stacked‑bar chart styled as a tornado diagram, assigns a unique solid‑fill color to each bar, recalculates the layout, exports the chart as a PNG image, and saves the workbook as XLSX.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sales data (positive and negative values to create a tornado effect)
            sheet.Cells["A1"].PutValue("Region");
            sheet.Cells["B1"].PutValue("Sales");

            string[] regions = { "North", "South", "East", "West" };
            int[] sales = { 120, -80, 150, -60 };

            for (int i = 0; i < regions.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(regions[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(sales[i]);    // Column B
            }

            // Add a stacked bar chart (tornado style) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart (including headers)
            chart.SetChartDataRange("A1:B5", true);

            // Apply custom colors to each data point
            Color[] pointColors = { Color.Red, Color.Green, Color.Blue, Color.Orange };
            for (int i = 0; i < chart.NSeries[0].Points.Count && i < pointColors.Length; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];
                point.Area.FillFormat.SolidFill.Color = pointColors[i];
            }

            // Recalculate the chart layout after modifications
            chart.Calculate();

            // Export the chart as a PNG image (default format is PNG)
            string pngPath = "tornado_chart.png";
            chart.ToImage(pngPath);
            Console.WriteLine($"Chart image saved to: {Path.GetFullPath(pngPath)}");

            // Save the workbook containing the chart
            string xlsxPath = "tornado_chart.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(xlsxPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
