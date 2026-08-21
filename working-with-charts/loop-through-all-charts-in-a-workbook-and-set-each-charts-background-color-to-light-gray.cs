// Title: Set Light Gray Opaque Background for All Charts in an Aspose.Cells Workbook (C#)
// Description: Learn how to loop through every worksheet and chart in an Aspose.Cells workbook and apply a light‑gray, opaque background to each chart's ChartArea using C#. The example creates a sample chart, updates the background color, sets the background mode to Opaque, and saves the workbook.
// Keywords: Aspose.Cells chart background color C# | set chart area fill Aspose.Cells | loop through charts workbook | light gray chart background | opaque background mode Aspose.Cells | C# Excel chart styling | global Aspose.Cells example | GitHub Aspose.Cells chart demo
// Common Searches: how to change chart background color in Aspose.Cells C# | apply same background to all charts in Excel using Aspose.Cells | make chart area opaque Aspose.Cells | set light gray fill for chart area programmatically | iterate over charts in workbook Aspose.Cells
// Developer Intent: Apply a uniform light‑gray, opaque background to every chart in a workbook with Aspose.Cells and C#.
// Use Cases: Standardize chart appearance across automatically generated financial reports. | Enforce corporate branding by giving all Excel charts a neutral gray background. | Prepare charts for high‑quality PDF or print output where a solid fill is required.
// AI Prompts: Write C# code that uses Aspose.Cells to set the ChartArea background color to a custom RGB value and make it opaque for all charts in a workbook. | Show how to skip pie charts while applying a light gray background to other chart types in Aspose.Cells. | Explain how to modify the loop to read the background color from a configuration file and apply it to each chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartBackgroundDemo
{
    // Learn how to loop through every worksheet and chart in an Aspose.Cells workbook and apply a light‑gray, opaque background to each chart's ChartArea using C#. The example creates a sample chart, updates the background color, sets the background mode to Opaque, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path) or create a new one
            Workbook workbook = new Workbook(); // creates a new workbook
            // If you need to load: Workbook workbook = new Workbook("input.xlsx");

            // Example: add a worksheet and a chart to demonstrate the operation
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Loop through all worksheets and their charts
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart ch in ws.Charts)
                {
                    // Set the background color of the chart area to light gray
                    ch.ChartArea.Area.BackgroundColor = Color.LightGray;

                    // Ensure the background is opaque so the color is visible
                    ch.ChartArea.BackgroundMode = BackgroundMode.Opaque;
                }
            }

            // Save the workbook
            workbook.Save("OutputWithChartBackground.xlsx");
        }
    }
}
