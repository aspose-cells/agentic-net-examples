// Title: Clone a Chart, Modify Its Secondary Axis, and Add to a New Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column chart with a secondary Y‑axis, clone the worksheet (including the chart), and then change the cloned chart's secondary axis title, minimum, maximum, and major unit before saving the file.
// Keywords: Aspose.Cells chart cloning | secondary axis modification | copy worksheet with chart C# | Aspose.Cells secondary value axis | clone chart Aspose.Cells .NET | chart style change Aspose.Cells | Worksheet.AddCopy example
// Common Searches: Aspose.Cells clone chart and edit secondary axis | C# copy worksheet with chart and change axis settings | How to modify secondary Y axis of a cloned chart in Aspose.Cells | Duplicate chart in new sheet Aspose.Cells .NET | Change secondary axis range after chart copy
// Developer Intent: The developer needs to duplicate an existing chart, adjust its secondary axis parameters, and place the modified chart on a separate worksheet without altering the original.
// Use Cases: Create a master chart template, clone it to multiple report sheets, and set unique secondary‑axis scales for each dataset. | Automate generation of dashboards where each worksheet shows a copy of the base chart with customized axis ranges for comparative analysis. | Maintain an unchanged source chart while producing styled clones with different secondary‑axis titles and ranges for presentation purposes.
// AI Prompts: Write C# code using Aspose.Cells to clone a chart from one worksheet to another and set a new title, minimum, maximum, and major unit for the secondary value axis. | Show how to duplicate a worksheet that contains a chart, then access the cloned chart and change its style and secondary axis properties. | Explain the steps for using Worksheet.AddCopy to copy a sheet with charts and subsequently modify the cloned chart's secondary axis in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    // Demonstrates how to create a workbook, add a column chart with a secondary Y‑axis, clone the worksheet (including the chart), and then change the cloned chart's secondary axis title, minimum, maximum, and major unit before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate sample data
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

            // Add a chart to the source sheet
            int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart originalChart = sourceSheet.Charts[chartIdx];

            // Set chart data
            originalChart.NSeries.Add("B2:B4", true);
            originalChart.NSeries.Add("C2:C4", true);
            originalChart.NSeries.CategoryData = "A2:A4";

            // Plot second series on secondary Y axis
            originalChart.NSeries[1].PlotOnSecondAxis = true;

            // Configure secondary value axis of the original chart
            Axis secAxis = originalChart.SecondValueAxis;
            secAxis.Title.Text = "Original Secondary Axis";
            secAxis.MinValue = 0;
            secAxis.MaxValue = 6000;
            secAxis.MajorUnit = 1000;

            // Clone the worksheet (which also clones the chart) into a new worksheet
            int clonedSheetIdx = workbook.Worksheets.AddCopy("Source");
            Worksheet clonedSheet = workbook.Worksheets[clonedSheetIdx];
            clonedSheet.Name = "Cloned";

            // Access the cloned chart (the first chart in the cloned sheet)
            Chart clonedChart = clonedSheet.Charts[0];

            // Modify secondary axis settings of the cloned chart
            Axis clonedSecAxis = clonedChart.SecondValueAxis;
            clonedSecAxis.Title.Text = "Cloned Secondary Axis";
            clonedSecAxis.MinValue = 500;   // new minimum
            clonedSecAxis.MaxValue = 5500;  // new maximum
            clonedSecAxis.MajorUnit = 500;  // new major unit

            // Optionally, change the chart style to differentiate it
            clonedChart.Style = 3;

            // Save the workbook
            workbook.Save("ChartCloneWithModifiedSecondaryAxis.xlsx");
        }
    }
}
