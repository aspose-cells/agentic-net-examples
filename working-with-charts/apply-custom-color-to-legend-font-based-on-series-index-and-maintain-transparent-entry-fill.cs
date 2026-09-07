// Title: How to set individual legend entry font colors by series index and keep legend backgrounds transparent using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates over chart series and assigns a unique font color to each legend entry, using a predefined color array and cycling when needed. | Show how to configure LegendEntry.BackgroundMode to Transparent for all series in an Aspose.Cells column chart while customizing the legend text colors. | Provide a complete Aspose.Cells example that creates a column chart, applies per‑series legend font colors, ensures transparent legend backgrounds, and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# change legend text color for each series in a column chart | set transparent background for legend entries using Aspose.Cells .NET | loop through NSeries to customize legend font colors Aspose.Cells example | apply different colors to chart legend entries based on series index Aspose.Cells | Aspose.Cells chart legend styling per series transparent background
// Tags: custom legend fonts Aspose.Cells | legend background transparency Aspose.Cells | series based legend colors C# | column chart legend formatting Aspose.Cells | legend color cycling Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook with sample data, adds a column chart, then iterates over each series to set the legend entry's font color from a predefined array (cycling as needed) and forces the legend entry background to be transparent before saving the file as CustomLegendColors.xlsx.
class CustomLegendColors
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
        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);
        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(80);
        sheet.Cells["C3"].PutValue(130);
        sheet.Cells["C4"].PutValue(170);
        sheet.Cells["D1"].PutValue("Series 3");
        sheet.Cells["D2"].PutValue(200);
        sheet.Cells["D3"].PutValue(210);
        sheet.Cells["D4"].PutValue(190);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add series to the chart
        chart.NSeries.Add("B2:B4", true); // Series 1
        chart.NSeries.Add("C2:C4", true); // Series 2
        chart.NSeries.Add("D2:D4", true); // Series 3
        chart.NSeries.CategoryData = "A2:A4";

        // Define a set of colors to apply to legend entries
        Color[] legendColors = new Color[]
        {
            Color.Red,
            Color.Green,
            Color.Blue,
            Color.Orange,
            Color.Purple
        };

        // Apply custom font color to each legend entry based on its series index
        for (int i = 0; i < chart.NSeries.Count; i++)
        {
            LegendEntry entry = chart.NSeries[i].LegendEntry;

            // Set the font color using the predefined color array (cycle if more series than colors)
            entry.Font.Color = legendColors[i % legendColors.Length];

            // Keep the legend entry background transparent
            entry.BackgroundMode = BackgroundMode.Transparent;
        }

        // Save the workbook
        workbook.Save("CustomLegendColors.xlsx");
    }
}
