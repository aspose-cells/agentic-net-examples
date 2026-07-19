// Title: Apply a Custom Color Palette to All Series in an Aspose.Cells Chart (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart with multiple series, and uses the SeriesCollection.ChangeColors method with ChartColorPaletteType.MonochromaticPalette1 to apply a uniform custom palette to every series before saving the file.
// Keywords: Aspose.Cells chart colors C# | ChangeColors method | ChartColorPaletteType | custom palette Aspose.Cells | multi‑series chart styling | programmatic Excel chart formatting | apply monochromatic palette | C# Excel chart color theme
// Common Searches: Aspose.Cells change all series colors C# | apply monochromatic palette to chart Aspose.Cells | SeriesCollection.ChangeColors example | set custom chart theme in .NET Excel | how to use ChartColorPaletteType in Aspose.Cells
// Developer Intent: Use Aspose.Cells for .NET to assign a single built‑in color palette to every series in a multi‑series chart.
// Use Cases: Generate Excel reports where chart series share a brand‑consistent color scheme. | Automate the styling of charts in bulk workbooks to match corporate visual guidelines. | Update existing charts programmatically to use a predefined palette without manually editing each series.
// AI Prompts: Show how to apply ChartColorPaletteType.VibrantPalette to all series of a chart in Aspose.Cells C#. | Provide code that iterates through each series and sets individual colors instead of using ChangeColors. | Explain how to create a custom color palette and assign it to a chart's SeriesCollection in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomPaletteDemo
{
    // Creates a workbook, adds sample data, inserts a column chart with multiple series, and uses the SeriesCollection.ChangeColors method with ChartColorPaletteType.MonochromaticPalette1 to apply a uniform custom palette to every series before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // Populate sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");

            ws.Cells["B1"].PutValue("Series1");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            ws.Cells["C1"].PutValue("Series2");
            ws.Cells["C2"].PutValue(15);
            ws.Cells["C3"].PutValue(25);
            ws.Cells["C4"].PutValue(35);

            // Add a column chart to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range for the series and categories
            chart.NSeries.Add("B1:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a monochromatic palette to all series in the chart
            // Using the ChangeColors method of SeriesCollection
            SeriesCollection seriesColl = chart.NSeries;
            seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette1);

            // Save the workbook
            workbook.Save("CustomPaletteChart.xlsx");
        }
    }
}
