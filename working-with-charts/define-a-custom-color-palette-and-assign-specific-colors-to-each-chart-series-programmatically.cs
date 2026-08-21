// Title: Define a Custom Color Palette and Set Chart Series Colors in Aspose.Cells (C#)
// Description: Creates a workbook, adds sample data for three series, inserts a column chart, replaces palette index 55 with a custom teal color, and assigns distinct foreground colors (red, teal, deep blue) to each series before saving the file as an XLSX workbook.
// Keywords: Aspose.Cells custom palette C# | change workbook palette Aspose.Cells | chart series color programmatically | column chart custom colors .NET | set series ForegroundColor Aspose | Aspose.Cells chart styling | C# Excel chart color palette
// Common Searches: how to change Aspose.Cells workbook palette | assign specific colors to chart series Aspose.Cells | custom teal color index 55 Aspose.Cells | set ForegroundColor for chart series C# | Aspose.Cells column chart custom colors
// Developer Intent: Programmatically define a custom color palette and apply individual colors to each series of a chart generated with Aspose.Cells for .NET.
// Use Cases: Generate a column chart that matches corporate brand colors by assigning red, teal, and blue to three product‑line series. | Replace a default palette entry with a corporate teal shade (index 55) and reuse it across multiple charts in a report. | Improve readability of financial dashboards by programmatically setting distinct series colors without manual Excel editing.
// AI Prompts: Show me C# code to add a custom palette entry at index 55 in Aspose.Cells and use it for a chart series. | Provide an Aspose.Cells example that sets different ForegroundColor values for each series of a column chart. | Explain step‑by‑step how to change the workbook palette and then assign those colors to chart series in .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsCustomPaletteDemo
{
    // Creates a workbook, adds sample data for three series, inserts a column chart, replaces palette index 55 with a custom teal color, and assigns distinct foreground colors (red, teal, deep blue) to each series before saving the file as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for three series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Series 1
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Series 2
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Series 3
            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series (B1:D4) and category data (A2:A4)
            chart.NSeries.Add("B1:D4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // OPTIONAL: Define a custom palette entry if you need colors not present in the default palette
            // Here we replace palette index 55 with a custom teal color
            Color customTeal = Color.FromArgb(0, 128, 128);
            workbook.ChangePalette(customTeal, 55);

            // Assign specific colors to each series using the Area.ForegroundColor property
            // Series 1 – use a bright red
            chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(255, 0, 0);
            // Series 2 – use the custom teal from the palette (index 55)
            chart.NSeries[1].Area.ForegroundColor = customTeal;
            // Series 3 – use a deep blue
            chart.NSeries[2].Area.ForegroundColor = Color.FromArgb(0, 0, 139);

            // Save the workbook
            workbook.Save("CustomPaletteChart.xlsx");
        }
    }
}
