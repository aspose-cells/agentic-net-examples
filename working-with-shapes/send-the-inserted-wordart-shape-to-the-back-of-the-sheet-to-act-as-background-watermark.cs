// Title: Add a WordArt background watermark in Excel with Aspose.Cells for .NET – send shape to back
// Description: C# example that creates a workbook, inserts a WordArt shape containing "CONFIDENTIAL", moves the shape behind all cells with ToFrontOrBack(1), applies 80% transparency, and saves the file as WordArtBackground.xlsx.
// Keywords: Aspose.Cells | C# WordArt watermark | send shape to back | ToFrontOrBack method | Excel background watermark .NET | Aspose.Cells shape ordering | WordArt transparency Aspose | add WordArt Aspose.Cells | programmatic Excel watermark | Aspose.Cells API
// Common Searches: How to place WordArt behind cells using Aspose.Cells | Aspose.Cells C# send shape to back | Create semi‑transparent Excel watermark with WordArt in .NET | Aspose.Cells ToFrontOrBack example | Programmatic Excel watermark Aspose.Cells
// Developer Intent: Insert a WordArt shape and position it behind worksheet content to serve as a background watermark.
// Use Cases: Generate confidential reports where a faint "CONFIDENTIAL" watermark appears behind all data. | Design branded templates by sending a logo WordArt shape to the back, providing a subtle background on each sheet. | Automate watermarking across multiple worksheets by adding WordArt, sending it to the back, and adjusting transparency.
// AI Prompts: Write C# code with Aspose.Cells that adds a WordArt shape, sends it to the back of the sheet, and sets 70% transparency for a watermark. | Explain the ToFrontOrBack method in Aspose.Cells and show how to use it to position shapes behind cells. | Provide a sample that applies a WordArt background watermark to every worksheet in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, inserts a WordArt shape containing "CONFIDENTIAL", moves the shape behind all cells with ToFrontOrBack(1), applies 80% transparency, and saves the file as WordArtBackground.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to demonstrate the watermark effect
            sheet.Cells["A1"].PutValue("Sample content behind watermark.");

            // Insert a WordArt shape that will serve as the watermark
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset style
                "CONFIDENTIAL",                  // watermark text
                0,   // top row index
                0,   // left column index
                0,   // vertical offset (pixels)
                0,   // horizontal offset (pixels)
                200, // height (pixels)
                600  // width (pixels)
            );

            // Send the WordArt shape to the back so it appears behind other content
            wordArt.ToFrontOrBack(1); // 1 = send to back

            // Make the watermark semi‑transparent (optional)
            wordArt.FillFormat.Transparency = 0.8; // 80% transparent

            // Save the workbook
            workbook.Save("WordArtBackground.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
