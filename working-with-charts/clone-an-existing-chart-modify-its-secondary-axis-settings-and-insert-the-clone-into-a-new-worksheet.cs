// Title: Clone a chart, adjust its secondary axis, and place the clone on a new worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that copies a worksheet containing a chart, retrieves the cloned chart, updates the secondary value axis title, minimum, maximum, and major unit, then saves the workbook. | Show how to duplicate a chart by cloning its worksheet with Aspose.Cells, modify the cloned chart's secondary axis properties, and export the result to an .xlsx file. | Provide a step‑by‑step example that creates a column chart with a secondary axis, clones the worksheet, changes the secondary axis settings of the cloned chart, and writes the workbook to disk.
// Common Searches: how to clone a chart and change secondary axis in Aspose.Cells C# | Aspose.Cells copy worksheet with chart and modify secondary value axis programmatically | C# example for duplicating a chart and setting secondary axis range using Aspose.Cells
// Tags: aspose.cells chart duplication | secondary axis configuration aspose.cells | worksheet copy with embedded chart aspose.cells | c# adjust secondary value axis | export workbook with modified chart aspose.cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneExample
{
    // The example creates a workbook, adds sample data, builds a column chart with a secondary value axis, clones the worksheet (including the chart), updates the cloned chart's secondary axis title and range, and saves the file as ClonedChartWithModifiedSecondaryAxis.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["A4"].PutValue("C");

            sourceSheet.Cells["B1"].PutValue("Series 1");
            sourceSheet.Cells["B2"].PutValue(100);
            sourceSheet.Cells["B3"].PutValue(200);
            sourceSheet.Cells["B4"].PutValue(300);

            sourceSheet.Cells["C1"].PutValue("Series 2");
            sourceSheet.Cells["C2"].PutValue(5000);
            sourceSheet.Cells["C3"].PutValue(3000);
            sourceSheet.Cells["C4"].PutValue(1000);

            // Add a chart to the source worksheet
            int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart sourceChart = sourceSheet.Charts[chartIdx];

            // Set chart data
            sourceChart.NSeries.Add("B2:B4", true);
            sourceChart.NSeries.Add("C2:C4", true);
            sourceChart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary value axis
            sourceChart.NSeries[1].PlotOnSecondAxis = true;

            // Configure the secondary value axis of the original chart (optional)
            Axis originalSecondAxis = sourceChart.SecondValueAxis;
            originalSecondAxis.Title.Text = "Original Secondary Axis";
            originalSecondAxis.MinValue = 0;
            originalSecondAxis.MaxValue = 6000;
            originalSecondAxis.MajorUnit = 1000;

            // -----------------------------------------------------------------
            // Clone the worksheet (which also clones the chart) using AddCopy
            // -----------------------------------------------------------------
            int clonedSheetIdx = workbook.Worksheets.AddCopy(0); // copy the first worksheet
            Worksheet clonedSheet = workbook.Worksheets[clonedSheetIdx];

            // Retrieve the cloned chart (it will have the same index as in the source sheet)
            Chart clonedChart = clonedSheet.Charts[chartIdx];

            // Modify secondary axis settings of the cloned chart
            Axis clonedSecondAxis = clonedChart.SecondValueAxis;
            clonedSecondAxis.Title.Text = "Cloned Secondary Axis";
            clonedSecondAxis.MinValue = 0;
            clonedSecondAxis.MaxValue = 5000;
            clonedSecondAxis.MajorUnit = 1000;

            // Save the workbook with the original and cloned charts
            workbook.Save("ClonedChartWithModifiedSecondaryAxis.xlsx");
        }
    }
}
