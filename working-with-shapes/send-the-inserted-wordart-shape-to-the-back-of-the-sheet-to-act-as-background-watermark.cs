// Title: How to place a WordArt shape behind cells as a background watermark in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set WordArt shape behind worksheet cells for watermark | C# Aspose.Cells add semi transparent WordArt watermark and send to back | How to use ZOrderPosition to move Excel shape to background with Aspose.Cells | Create background watermark with WordArt in .xlsx using Aspose.Cells .NET | Aspose.Cells example for sending WordArt shape to back of sheet
// Tags: WordArt shape background watermark Aspose.Cells | ZOrderPosition shape placement Aspose.Cells .NET | semi transparent WordArt watermark Excel | move Excel shape behind cells Aspose.Cells | add WordArt watermark to .xlsx using C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtBackgroundWatermark
{
    // The example creates a new workbook, inserts a WordArt shape with custom text, color, size, and transparency, sets its ZOrderPosition to 0 so it sits behind all cells, and saves the worksheet as an XLSX file, effectively using the shape as a background watermark.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a WordArt shape that will be used as a watermark
                // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
                Shape wordArt = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1, // preset style
                    "CONFIDENTIAL",                  // watermark text
                    5,                               // top row index
                    0,                               // top offset in pixels
                    2,                               // left column index
                    0,                               // left offset in pixels
                    100,                             // height in pixels
                    600);                            // width in pixels

                // Customize appearance (make it semi‑transparent)
                wordArt.Fill.Transparency = 0.7;               // 70% transparent
                wordArt.HasLine = false;                       // hide border line
                wordArt.Font.Color = System.Drawing.Color.Red; // text color
                wordArt.Font.Size = 48;                        // font size
                wordArt.Font.IsBold = true;                    // bold text
                wordArt.Font.Underline = FontUnderlineType.Single;

                // Send the WordArt shape to the back so that it behaves as a background watermark
                wordArt.ZOrderPosition = 0;

                // Save the workbook
                string outputPath = "WordArtWatermark.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The WordArt shape has been sent to the back of the sheet as a background watermark.");
            }
            catch (Exception ex)
            {
                // Report any errors that occur during processing
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
