// Title: Aspose.Cells C# – Set Primary Y‑Axis Major Unit to 10 in a Column Chart
// Description: Creates a workbook, adds sample data, inserts a column chart, disables automatic major unit on the primary Y‑axis, sets MajorUnit = 10, and saves the file as ChartWithCustomMajorUnit.xlsx.
// Keywords: Aspose.Cells Y axis major unit | C# chart major grid interval | disable automatic major unit Aspose | primary value axis custom interval | column chart grid spacing 10
// Common Searches: Aspose.Cells set primary Y axis major unit C# | C# chart ValueAxis IsAutomaticMajorUnit false example | how to fix Y‑axis grid interval in Aspose.Cells | custom major unit for Excel chart using Aspose.Cells
// Developer Intent: Configure the primary Y‑axis of an Aspose.Cells chart to use a fixed major unit of 10, overriding the automatic scaling.
// Use Cases: Produce financial dashboards where Y‑axis ticks must increase by 10 for uniform visual comparison. | Standardize grid lines across multiple charts in a single workbook for corporate reporting templates. | Generate regulatory‑compliant Excel reports that require explicit Y‑axis scaling.
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart and sets the primary Y‑axis major unit to 5. | Write a script that iterates through all charts in a workbook and applies a fixed major unit of 10 to each ValueAxis while disabling automatic scaling.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, disables automatic major unit on the primary Y‑axis, sets MajorUnit = 10, and saves the file as ChartWithCustomMajorUnit.xlsx.
    class SetPrimaryYAxisMajorUnit
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(5);
            worksheet.Cells["B3"].PutValue(15);
            worksheet.Cells["B4"].PutValue(25);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Disable automatic major unit and set the major unit interval to 10 on the primary Y axis
            chart.ValueAxis.IsAutomaticMajorUnit = false;
            chart.ValueAxis.MajorUnit = 10;

            // Save the workbook
            workbook.Save("ChartWithCustomMajorUnit.xlsx");
        }
    }
}
