// Title: Clone a chart, reassign its data range, and insert it on a different worksheet using Aspose.Cells for .NET (C#)
// Description: A C# walkthrough that creates a workbook, populates source data, builds a column chart, copies the chart to another sheet, points the copy to a new data block, and saves the file. The sample highlights chart duplication, series rebinding, and cross‑worksheet placement with Aspose.Cells.
// Keywords: Aspose.Cells chart clone C# | duplicate chart Aspose.Cells | change chart data source .NET | copy chart to another worksheet | Aspose.Cells series range update | C# Excel chart example | Aspose.Cells chart manipulation
// Common Searches: how to copy a chart to a different sheet with Aspose.Cells | Aspose.Cells C# change chart data range after cloning | duplicate Excel chart programmatically using Aspose | rebind chart series to new cells in Aspose.Cells | clone and move chart between worksheets in .NET
// Developer Intent: The developer needs to replicate an existing chart, bind it to a new data set, and place the replica on another worksheet.
// Use Cases: Create a template chart once and reuse it across multiple department sheets with their own data. | Generate comparative dashboards by cloning a baseline chart and linking each copy to forecast versus actual values. | Automate report generation where each region gets a personalized chart derived from a master design.
// AI Prompts: Show C# code that clones an Aspose.Cells chart and updates its series to a different cell range. | Explain how to preserve chart formatting while changing the data source after copying the chart to another worksheet. | Provide a script to duplicate several charts and assign each a unique data range on separate sheets using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    // A C# walkthrough that creates a workbook, populates source data, builds a column chart, copies the chart to another sheet, points the copy to a new data block, and saves the file. The sample highlights chart duplication, series rebinding, and cross‑worksheet placement with Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a workbook and add a source worksheet
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Populate sample data for the source chart
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Value");
                sourceSheet.Cells["A2"].PutValue("A");
                sourceSheet.Cells["B2"].PutValue(10);
                sourceSheet.Cells["A3"].PutValue("B");
                sourceSheet.Cells["B3"].PutValue(20);
                sourceSheet.Cells["A4"].PutValue("C");
                sourceSheet.Cells["B4"].PutValue(30);

                // Add a chart to the source worksheet
                int srcChartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart srcChart = sourceSheet.Charts[srcChartIdx];
                srcChart.NSeries.Add("B2:B4", true);
                srcChart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // 2. Add a destination worksheet where the clone will be placed
                // -------------------------------------------------
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // -------------------------------------------------
                // 3. Clone the chart: create a new chart with the same type
                // -------------------------------------------------
                int clonedChartIdx = destSheet.Charts.Add(srcChart.Type, 5, 0, 15, 5);
                Chart clonedChart = destSheet.Charts[clonedChartIdx];

                // -------------------------------------------------
                // 4. Modify the data source of the cloned chart
                //    (for demonstration, use a different range on the destination sheet)
                // -------------------------------------------------
                // Populate new data on the destination sheet
                destSheet.Cells["C1"].PutValue("Category");
                destSheet.Cells["D1"].PutValue("Value");
                destSheet.Cells["C2"].PutValue("X");
                destSheet.Cells["D2"].PutValue(40);
                destSheet.Cells["C3"].PutValue("Y");
                destSheet.Cells["D3"].PutValue(50);
                destSheet.Cells["C4"].PutValue("Z");
                destSheet.Cells["D4"].PutValue(60);

                // Set the new data range for the cloned chart
                clonedChart.NSeries.Add("D2:D4", true);
                clonedChart.NSeries.CategoryData = "C2:C4";

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                workbook.Save("ChartCloneResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
