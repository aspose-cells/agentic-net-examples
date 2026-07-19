// Title: Insert a Rotated WordArt Watermark into an Aspose.Cells Chart (C#)
// Description: Demonstrates how to create a workbook, add a column chart, and overlay a semi‑transparent WordArt shape as a diagonal watermark using Aspose.Cells' AddTextEffectInChart method. The shape is set to 70% transparency and rotated 30°, then the file is saved as an XLSX document.
// Keywords: Aspose.Cells WordArt watermark | rotate WordArt chart C# | AddTextEffectInChart example | .NET chart watermark transparency | Aspose.Cells diagonal watermark | GitHub Aspose.Cells chart sample | C# Excel chart shape rotation
// Common Searches: how to add WordArt watermark to Aspose.Cells chart | rotate chart shape 30 degrees Aspose.Cells C# | set transparency for WordArt in Excel chart using Aspose | Aspose.Cells AddTextEffectInChart usage | GitHub example of chart watermark with Aspose.Cells
// Developer Intent: Overlay a semi‑transparent, diagonal WordArt watermark on a chart programmatically.
// Use Cases: Mark confidential or draft charts with a clear diagonal label. | Brand internal reports by placing a company name as a subtle background on each chart. | Comply with regulatory requirements by adding a visible watermark to exported Excel charts.
// AI Prompts: Write C# code that adds a 45° rotated WordArt watermark with 80% transparency to a pie chart using Aspose.Cells. | Explain how to control the position, size, and rotation of a WordArt shape inside an Aspose.Cells chart. | Provide a workaround for hiding the outline of a WordArt shape when the Line.IsVisible property is unavailable.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart, and overlay a semi‑transparent WordArt shape as a diagonal watermark using Aspose.Cells' AddTextEffectInChart method. The shape is set to 70% transparency and rotated 30°, then the file is saved as an XLSX document.
    public class WordArtWatermarkInChart
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart covering the data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Insert WordArt into the chart as a watermark
            // Units for position and size are 1/4000 of the chart area
            Shape wordArt = chart.Shapes.AddTextEffectInChart(
                MsoPresetTextEffect.TextEffect2,   // preset effect
                "CONFIDENTIAL",                    // watermark text
                "Arial Black",                     // font name
                48,                                // font size
                true,                              // bold
                false,                             // italic
                1000,                              // top offset
                1000,                              // left offset
                2000,                              // height
                2000);                             // width

            // Make the WordArt appear as a semi‑transparent watermark
            wordArt.Fill.Transparency = 0.7;   // 70% transparent

            // Hide the outline of the WordArt (Line property does not expose IsVisible in this version)
            // wordArt.Line.IsVisible = false; // Removed due to API limitation

            // Rotate the WordArt 30 degrees for a diagonal effect
            wordArt.RotationAngle = 30;

            // Save the workbook with the chart watermark
            string outputPath = "ChartWordArtWatermark.xlsx";
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
    }
}
