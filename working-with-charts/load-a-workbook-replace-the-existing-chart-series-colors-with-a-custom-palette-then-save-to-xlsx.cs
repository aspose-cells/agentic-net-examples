// Title: Change Excel Chart Series Colors with a Monochromatic Palette using Aspose.Cells for .NET
// Description: Loads an existing XLSX workbook, creates a column chart if none exists, accesses the first chart's series collection, applies the MonochromaticPalette5 scheme via the ChangeColors method, and saves the updated file.
// Keywords: Aspose.Cells C# chart colors | ChangeColors method | ChartColorPaletteType MonochromaticPalette5 | replace Excel chart series colors | programmatic chart palette | Excel automation Aspose | update chart series colors .NET | custom chart color scheme | Excel chart styling code | Aspose.Cells chart example
// Common Searches: how to change chart series colors with Aspose.Cells | Aspose.Cells ChangeColors example C# | apply monochromatic palette to Excel chart programmatically | replace Excel chart colors in .NET | Aspose.Cells chart color palette types
// Developer Intent: Apply a predefined monochromatic palette to all series of an existing Excel chart and save the workbook.
// Use Cases: Standardize chart colors across reports to match corporate branding. | Automatically recolor charts when generating workbooks from templates. | Update legacy workbooks that use default palettes without manual editing.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, changes the chart series colors to a specified palette, and saves the workbook. | Show how to use the ChangeColors method with ChartColorPaletteType to apply MonochromaticPalette5 to a chart's series collection. | Explain how to create a chart only when none exists, then apply a custom color palette using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX workbook, creates a column chart if none exists, accesses the first chart's series collection, applies the MonochromaticPalette5 scheme via the ChangeColors method, and saves the updated file.
class ReplaceChartSeriesColors
{
    static void Main()
    {
        // Load the existing workbook from disk
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet ws = workbook.Worksheets[0];

        // If the worksheet has no chart, create a sample one (optional)
        if (ws.Charts.Count == 0)
        {
            // Sample data for the chart
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

            // Add a column chart and bind data
            int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B1:C4", true);
            chart.NSeries.CategoryData = "A2:A4";
        }

        // Retrieve the first chart in the worksheet
        Chart firstChart = ws.Charts[0];

        // Get the series collection of the chart
        SeriesCollection seriesColl = firstChart.NSeries;

        // Replace the existing series colors with a custom monochromatic palette
        // (choose any palette type you prefer)
        seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette5);

        // Save the workbook with the updated chart colors
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
