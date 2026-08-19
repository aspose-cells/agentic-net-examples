// Title: Aspose.Cells C# – Set Custom RGB Colors for Each Series in a Column Chart
// Description: Shows how to build a workbook, fill it with category and series data, insert a column chart, and assign different RGB values to individual series using the Aspose.Cells for .NET API before saving the file as XLSX.
// Keywords: Aspose.Cells chart series color | C# custom chart palette | RGB series styling Aspose.Cells | column chart series color .NET | Aspose.Cells chart customization | set series foreground color | chart series color array | Aspose.Cells SaveFormat Xlsx
// Common Searches: C# Aspose.Cells change series color | apply custom colors to chart series Aspose.Cells | set RGB values for chart series .NET | custom color palette for Aspose.Cells column chart | Aspose.Cells chart series styling example
// Developer Intent: Assign distinct RGB color values to each data series of a chart created with Aspose.Cells for .NET.
// Use Cases: Produce a sales chart where each product line matches its brand hue. | Build a financial dashboard with separate colors for revenue, expenses, and profit series. | Export a presentation‑ready chart that adheres to a corporate color scheme across all series.
// AI Prompts: Generate C# code with Aspose.Cells that applies custom RGB colors to each series of a line chart. | Explain how to replace the hard‑coded Color array with a reusable ColorPalette object in the example. | Show how to set both the fill and border colors for chart series using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, fill it with category and series data, insert a column chart, and assign different RGB values to individual series using the Aspose.Cells for .NET API before saving the file as XLSX.
class CustomChartPalette
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];

        // Populate data for the chart
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

        ws.Cells["D1"].PutValue("Series3");
        ws.Cells["D2"].PutValue(12);
        ws.Cells["D3"].PutValue(22);
        ws.Cells["D4"].PutValue(32);

        // Add a column chart
        int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 12);
        Chart chart = ws.Charts[chartIdx];

        // Set the data range for the series (B1:D4) and categories (A2:A4)
        chart.NSeries.Add("B1:D4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Define custom RGB colors for each series
        Color[] customColors = new Color[]
        {
            Color.FromArgb(79, 129, 189),   // Series1
            Color.FromArgb(192, 80, 77),   // Series2
            Color.FromArgb(255, 192, 0)    // Series3
        };

        // Apply the custom colors to the series areas
        for (int i = 0; i < chart.NSeries.Count && i < customColors.Length; i++)
        {
            chart.NSeries[i].Area.ForegroundColor = customColors[i];
        }

        // Save the workbook with the customized chart
        workbook.Save("CustomPaletteChart.xlsx", SaveFormat.Xlsx);
    }
}
