// Title: Aspose.Cells .NET: Create a Fully Transparent Chart Legend (C#)
// Description: Demonstrates how to generate a workbook, add a column chart, and make the legend completely transparent by setting Legend.BackgroundMode and each LegendEntry's BackgroundMode to Transparent and disabling text fill with IsTextNoFill. The result is a clear legend that does not obscure the worksheet background.
// Keywords: Aspose.Cells transparent legend | chart legend background mode transparent | remove legend fill Aspose.Cells | LegendEntry IsTextNoFill C# | Aspose.Cells chart styling | Excel chart invisible legend .NET
// Common Searches: how to make chart legend transparent in Aspose.Cells | Aspose.Cells remove fill from legend entries | set legend background mode to transparent C# | make Excel chart legend invisible using Aspose.Cells | transparent legend for column chart Aspose.Cells
// Developer Intent: Apply transparent styling to a chart legend and its entries so the legend becomes invisible while preserving the series data.
// Use Cases: Embedding a chart in a presentation where the legend must not cover underlying graphics. | Generating Excel reports with custom UI designs that require an invisible legend. | Preparing printable charts where the legend should be omitted without altering layout.
// AI Prompts: Show C# code to set a chart legend and its entries to fully transparent in Aspose.Cells. | Explain which Aspose.Cells properties hide a legend while keeping the chart series visible. | Provide step‑by‑step instructions for making an Excel chart legend invisible using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsLegendTransparentDemo
{
    // Demonstrates how to generate a workbook, add a column chart, and make the legend completely transparent by setting Legend.BackgroundMode and each LegendEntry's BackgroundMode to Transparent and disabling text fill with IsTextNoFill. The result is a clear legend that does not obscure the worksheet background.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(70);
            sheet.Cells["B4"].PutValue(55);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Make the whole legend background transparent
            chart.Legend.BackgroundMode = BackgroundMode.Transparent;

            // Ensure each legend entry has no fill (transparent background and no text fill)
            LegendEntryCollection entries = chart.Legend.LegendEntries;
            foreach (LegendEntry entry in entries)
            {
                // Transparent background for the entry
                entry.BackgroundMode = BackgroundMode.Transparent;

                // No fill for the text inside the entry
                entry.IsTextNoFill = true;
            }

            // Save the workbook
            workbook.Save("LegendTransparentDemo.xlsx");
        }
    }
}
