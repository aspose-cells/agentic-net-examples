// Title: C# – Resize an Aspose.Cells chart to span 10 rows × 15 columns using Height & Width
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, and resize the chart so it covers ten rows and fifteen columns. The example calculates pixel dimensions from the default row height (~20 px) and column width (~64 px) and sets ChartObject.Height, ChartObject.Width, UpperLeftRow, and UpperLeftColumn before saving the file.
// Keywords: Aspose.Cells chart resize C# | ChartObject Height Width | set chart size Aspose.Cells .NET | resize chart to cell range | pixel dimensions Excel chart | programmatic chart scaling | Aspose.Cells chart positioning
// Common Searches: Aspose.Cells set chart height and width | Resize chart to specific rows and columns Aspose.Cells | Calculate chart pixel size from row height column width | C# Aspose.Cells adjust chart dimensions | How to programmatically resize Excel chart with Aspose
// Developer Intent: Adjust a chart’s Height and Width so it occupies exactly ten worksheet rows and fifteen columns.
// Use Cases: Generate reports where charts must align with a predefined cell block for a uniform layout. | Automate chart scaling when exporting data to Excel to match page‑margin constraints. | Re‑size and reposition charts after data refresh to keep visual consistency with surrounding cells.
// AI Prompts: Write C# code with Aspose.Cells that resizes a chart to cover 10 rows and 15 columns, computing pixel values from default row height and column width. | Explain how to derive ChartObject.Height and ChartObject.Width based on worksheet row height and column width in Aspose.Cells. | Provide a snippet that moves and resizes an existing Aspose.Cells chart to fit a specific cell range using .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartResize
{
    // Demonstrates how to create a workbook, add sample data, insert a column chart, and resize the chart so it covers ten rows and fifteen columns. The example calculates pixel dimensions from the default row height (~20 px) and column width (~64 px) and sets ChartObject.Height, ChartObject.Width, UpperLeftRow, and UpperLeftColumn before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ------------------------------------------------------------
            // Resize the chart to occupy ten rows and fifteen columns.
            // Aspose.Cells allows setting the chart size in pixels via the
            // ChartObject's Width and Height properties. To approximate the
            // size of ten rows and fifteen columns, we multiply the default
            // row height (15 points ≈ 20 pixels) and default column width
            // (64 pixels) by the desired counts.
            // ------------------------------------------------------------
            const int defaultRowHeightPixels = 20;   // Approximate pixel height of a row
            const int defaultColumnWidthPixels = 64; // Approximate pixel width of a column

            // Set height to cover ten rows
            chart.ChartObject.Height = defaultRowHeightPixels * 10; // 200 pixels

            // Set width to cover fifteen columns
            chart.ChartObject.Width = defaultColumnWidthPixels * 15; // 960 pixels

            // Optionally, reposition the chart to start at a specific cell (e.g., row 5, column 0)
            chart.ChartObject.UpperLeftRow = 5;
            chart.ChartObject.UpperLeftColumn = 0;

            // Save the workbook (save rule)
            workbook.Save("ResizedChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
