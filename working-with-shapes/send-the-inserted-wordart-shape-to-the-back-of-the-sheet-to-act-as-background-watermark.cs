// Title: Insert a WordArt text effect as a background watermark and send it behind all objects in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a WordArt shape with the text "CONFIDENTIAL" to the first worksheet, sets its ZOrderPosition to 0, rotates it -45°, applies 80% transparency, and saves the workbook as an .xlsx file. | Write a C# snippet using Aspose.Cells to create a WordArt watermark, place it behind existing cells, adjust its rotation and fill transparency, and export the result to Watermark.xlsx.
// Common Searches: aspnet add wordart watermark to excel worksheet using aspose.cells | c# send shape to back layer in aspose.cells workbook | how to set zorderposition for wordart shape in aspose.cells | rotate wordart shape and set transparency in excel file with aspose.cells | create background watermark with text effect in aspose.cells c#
// Tags: add wordart shape as watermark Aspose.Cells | set shape ZOrderPosition Aspose.Cells | rotate wordart shape Aspose.Cells | apply transparency to shape Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, inserts a WordArt text effect shape containing "CONFIDENTIAL" on the first worksheet, sends the shape to the back by setting ZOrderPosition to 0, rotates it -45°, makes it 80% transparent, and saves the file as Watermark.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a WordArt shape (text effect) to be used as a watermark
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset text effect
                "CONFIDENTIAL",                  // watermark text
                "Arial",                         // font name
                72,                              // font size
                false,                           // bold
                false,                           // italic
                0, 0,                            // upper‑left cell (row, column)
                0, 0,                            // lower‑right cell (row, column)
                200, 500);                       // height and width

            // Send the shape to the back so it appears behind all other objects
            wordArt.ZOrderPosition = 0;

            // Optional: rotate and make the watermark semi‑transparent
            wordArt.RotationAngle = -45;
            wordArt.Fill.Transparency = 0.8;

            // Save the workbook
            workbook.Save("Watermark.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
