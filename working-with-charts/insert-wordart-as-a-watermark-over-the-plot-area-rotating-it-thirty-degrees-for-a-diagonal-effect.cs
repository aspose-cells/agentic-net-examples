// Title: Insert a diagonal WordArt watermark over a chart’s plot area using Aspose.Cells for .NET (C#)
// Common Searches: how to add a WordArt watermark to a chart plot area with Aspose.Cells C# | rotate WordArt shape 30 degrees in an Aspose.Cells chart | set transparency for WordArt watermark over a column chart using Aspose.Cells | Aspose.Cells example to overlay CONFIDENTIAL text on a chart | C# code to place WordArt over chart object in Aspose.Cells workbook
// Tags: Aspose.Cells add WordArt watermark to chart | rotate WordArt shape Aspose.Cells C# | WordArt transparency setting Aspose.Cells | chart plot area overlay Aspose.Cells | column chart WordArt overlay .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace WordArtWatermarkDemo
{
    // The sample creates a workbook, fills it with sample data, adds a column chart, and then inserts a WordArt shape that covers the chart's plot area as a watermark. The WordArt text "CONFIDENTIAL" is rotated 30° for a diagonal effect, its fill transparency is set to 60 %, and it is brought to the front before saving the file as WordArtWatermark.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate some sample data for the chart
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // -------------------------------------------------
                // 2. Add a column chart to the worksheet
                // -------------------------------------------------
                // Parameters: type, upper left row, upper left column, lower right row, lower right column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Quarterly Sales";

                // -------------------------------------------------
                // 3. Insert WordArt as a watermark over the plot area
                // -------------------------------------------------
                // Use the ChartObject to obtain position and size of the chart
                int topRow = chart.ChartObject.UpperLeftRow;      // Upper left row of the chart
                int leftColumn = chart.ChartObject.UpperLeftColumn; // Upper left column of the chart
                int chartHeight = chart.ChartObject.Height;       // Height in pixels
                int chartWidth = chart.ChartObject.Width;         // Width in pixels

                // Add WordArt covering the whole chart area
                Shape wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,   // Preset style
                    "CONFIDENTIAL",                     // Watermark text
                    topRow, 0,                          // Row index and vertical offset (pixels)
                    leftColumn, 0,                      // Column index and horizontal offset (pixels)
                    chartHeight, chartWidth);           // Size in pixels

                // Rotate the WordArt 30 degrees for a diagonal effect
                wordArt.RotationAngle = 30;

                // Make the watermark semi‑transparent so the chart remains readable
                wordArt.Fill.Transparency = 0.6; // 60 % transparent

                // Bring the WordArt to the front
                wordArt.ZOrderPosition = 0; // 0 = frontmost

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                string outputPath = "WordArtWatermark.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
