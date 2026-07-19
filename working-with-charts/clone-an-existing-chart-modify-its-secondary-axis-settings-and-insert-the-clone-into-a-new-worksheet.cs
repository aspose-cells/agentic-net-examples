// Title: Clone a Chart and Modify Its Secondary Axis on a New Worksheet – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook with a column chart that uses a secondary value axis, clone the worksheet (including the chart) with Worksheet.AddCopy, access the cloned chart, and change its secondary axis title, minimum, maximum and major unit before saving the file.
// Keywords: Aspose.Cells chart clone | secondary axis C# | Worksheet.AddCopy example | copy worksheet with chart .NET | modify secondary value axis Aspose.Cells | clone chart to new sheet | C# Aspose.Cells chart manipulation
// Common Searches: Aspose.Cells clone chart and edit secondary axis | Copy worksheet with chart C# Aspose.Cells | How to change secondary axis of a duplicated chart | Worksheet.AddCopy chart example | C# code to modify secondary value axis after cloning
// Developer Intent: The developer needs to duplicate an existing chart onto another worksheet and adjust the secondary axis properties of the copied chart using Aspose.Cells for .NET.
// Use Cases: Generate multiple reports from a template workbook where each report requires a different secondary‑axis scale. | Localize dashboards by cloning a base chart and customizing the secondary axis title and range for each region. | Build a multi‑sheet analytical workbook that reuses a common chart layout while showing distinct secondary‑axis values per sheet.
// AI Prompts: Write C# code with Aspose.Cells that clones a worksheet containing a chart, then updates the cloned chart's secondary value axis title, min, max, and major unit before saving. | Explain step‑by‑step how to use Worksheet.AddCopy to duplicate a chart and modify its secondary axis properties in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneExample
{
    // Demonstrates how to create a workbook with a column chart that uses a secondary value axis, clone the worksheet (including the chart) with Worksheet.AddCopy, access the cloned chart, and change its secondary axis title, minimum, maximum and major unit before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";

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

            // Add a chart to the source worksheet
            int chartIndex = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart sourceChart = sourceSheet.Charts[chartIndex];

            // Set chart data
            sourceChart.NSeries.Add("B2:B4", true);
            sourceChart.NSeries.Add("C2:C4", true);
            sourceChart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary value axis
            sourceChart.NSeries[1].PlotOnSecondAxis = true;

            // Configure the secondary value axis of the source chart
            Axis sourceSecondAxis = sourceChart.SecondValueAxis;
            sourceSecondAxis.Title.Text = "Original Secondary Axis";
            sourceSecondAxis.MinValue = 0;
            sourceSecondAxis.MaxValue = 6000;
            sourceSecondAxis.MajorUnit = 1000;

            // Clone the worksheet (including the chart) using AddCopy
            int clonedSheetIndex = workbook.Worksheets.AddCopy(0);
            Worksheet clonedSheet = workbook.Worksheets[clonedSheetIndex];
            clonedSheet.Name = "ClonedSheet";

            // Access the cloned chart (same index as in the source sheet)
            Chart clonedChart = clonedSheet.Charts[chartIndex];

            // Modify the secondary axis settings of the cloned chart
            Axis clonedSecondAxis = clonedChart.SecondValueAxis;
            clonedSecondAxis.Title.Text = "Cloned Secondary Axis";
            clonedSecondAxis.MinValue = 100;
            clonedSecondAxis.MaxValue = 5000;
            clonedSecondAxis.MajorUnit = 500;

            // Save the workbook
            workbook.Save("ClonedChartWorkbook.xlsx");
        }
    }
}
