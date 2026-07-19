// Title: Clone a Chart, Change Its Legend Position, and Save Both Charts with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column chart, duplicate the chart by reusing its type and data range, move the duplicated chart's legend to the bottom, disable legend overlay, and save the file containing both charts using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart duplication | clone chart Aspose.Cells | chart legend position | LegendPositionType.Bottom | GetChartDataRange | Charts.Add overload | Excel workbook chart copy | disable legend overlay | duplicate chart example
// Common Searches: Aspose.Cells duplicate chart C# example | how to copy a chart and change legend position in .NET | Aspose.Cells clone chart and set legend to bottom | C# add second chart with same data range Aspose.Cells | remove legend overlay for duplicated chart Aspose.Cells
// Developer Intent: Copy an existing chart, adjust the legend placement of the copy, and persist both charts in the same workbook.
// Use Cases: Create side‑by‑side visual comparisons by showing the original chart and a cloned chart with a different legend layout. | Automate report generation where a chart is reused with alternative legend positioning for printed or PDF output. | Apply multiple styling variations—such as legend location, overlay, or colors—by cloning a chart and customizing each instance before saving.
// AI Prompts: Generate C# code with Aspose.Cells that clones a chart, moves the cloned chart's legend to the right, and saves the workbook. | Show how to duplicate a chart, turn off legend overlay, add a title to the new chart, and export the file using Aspose.Cells for .NET. | Explain the steps to retrieve a chart's data range with GetChartDataRange and reuse it when adding another chart in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDuplication
{
    // Demonstrates how to create a workbook, add a column chart, duplicate the chart by reusing its type and data range, move the duplicated chart's legend to the bottom, disable legend overlay, and save the file containing both charts using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruits");
            sheet.Cells["A3"].PutValue("Vegetables");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);

            // -----------------------------------------------------------------
            // Add the original chart
            // -----------------------------------------------------------------
            int originalChartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart originalChart = sheet.Charts[originalChartIndex];
            originalChart.NSeries.Add("B2:B3", true);          // Values
            originalChart.NSeries.CategoryData = "A2:A3";     // Categories

            // -----------------------------------------------------------------
            // Duplicate the original chart
            // -----------------------------------------------------------------
            // Retrieve the chart type and data range of the original chart
            ChartType chartType = originalChart.Type;
            string dataRange = originalChart.GetChartDataRange(); // e.g., "A1:B3"
            // Add a new chart with the same type and data range.
            // The Add method with (ChartType, string, bool, int, int, int, int) allows specifying the data range.
            int duplicateChartIndex = sheet.Charts.Add(chartType, dataRange, true, 22, 0, 37, 8);
            Chart duplicateChart = sheet.Charts[duplicateChartIndex];

            // -----------------------------------------------------------------
            // Modify the legend position of the duplicated chart
            // -----------------------------------------------------------------
            duplicateChart.Legend.Position = LegendPositionType.Bottom; // Move legend to bottom
            // Optionally, adjust overlay behavior
            duplicateChart.Legend.IsOverLay = false;

            // -----------------------------------------------------------------
            // Save the workbook containing both charts
            // -----------------------------------------------------------------
            workbook.Save("DuplicatedChartDemo.xlsx");
        }
    }
}
