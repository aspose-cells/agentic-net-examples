// Title: Add Diagonal WordArt Watermark to an Aspose.Cells Chart (C#)
// Description: This example creates a workbook, fills it with sample data, adds a column chart, and inserts a WordArt text effect that spans the entire plot area. The WordArt is rotated 30° for a diagonal watermark, set to 70% transparency, and saved as ChartWithWordArtWatermark.xlsx.
// Keywords: Aspose.Cells chart watermark | C# WordArt watermark Aspose.Cells | rotate WordArt chart Aspose.Cells | transparent chart watermark C# | add text effect to chart Aspose.Cells | cover plot area with WordArt | AddTextEffectInChart Aspose.Cells
// Common Searches: Aspose.Cells add WordArt watermark to chart | C# rotate WordArt inside Aspose.Cells chart | set transparency for chart watermark Aspose.Cells | fit WordArt to chart plot area | diagonal chart watermark example .NET
// Developer Intent: Insert a semi‑transparent, 30° rotated WordArt shape that fully covers a chart’s plot area as a watermark.
// Use Cases: Mark financial or legal charts as CONFIDENTIAL with a diagonal watermark. | Brand internal presentations by overlaying a company name or slogan on charts. | Create draft versions of charts for review before final publishing.
// AI Prompts: Write C# code using Aspose.Cells to insert a WordArt text effect as a diagonal watermark on any chart type, with customizable text, font, size, rotation, and transparency. | Explain how to calculate the shape dimensions needed to fully cover the plot area of a chart in Aspose.Cells. | Show how to adjust the rotation angle and transparency of a WordArt shape after adding it to a chart with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsWatermarkExample
{
    // This example creates a workbook, fills it with sample data, adds a column chart, and inserts a WordArt text effect that spans the entire plot area. The WordArt is rotated 30° for a diagonal watermark, set to 70% transparency, and saved as ChartWithWordArtWatermark.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart covering rows 5‑15 and columns 0‑5
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Insert WordArt (text effect) into the chart.
            // The position and size are expressed in 1/4000 of the chart area.
            // Setting top, left, height, width to cover the whole plot area.
            Shape wordArt = chart.Shapes.AddTextEffectInChart(
                MsoPresetTextEffect.TextEffect2,   // preset effect
                "CONFIDENTIAL",                    // text
                "Arial Black",                     // font name
                48,                                // font size
                true,                              // bold
                false,                             // italic
                0,    // top offset (0/4000)
                0,    // left offset (0/4000)
                4000, // height (full chart height)
                4000  // width  (full chart width)
            );

            // Rotate the WordArt 30 degrees to achieve a diagonal watermark effect
            wordArt.RotationAngle = 30;

            // Make the WordArt semi‑transparent so the chart remains visible
            wordArt.FillFormat.Transparency = 0.7; // 70% transparent
            wordArt.LineFormat.IsVisible = false; // hide outline

            // Save the workbook with the watermarked chart
            workbook.Save("ChartWithWordArtWatermark.xlsx");
        }
    }
}
