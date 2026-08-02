// Title: Add WordArt as a Background Watermark in Excel using Aspose.Cells for .NET
// Description: C# example that creates a workbook, inserts a WordArt shape with custom text, sets its ZOrderPosition to 0 so it sits behind all other objects, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells WordArt watermark | C# Excel background shape | ZOrderPosition Aspose.Cells | send shape to back Excel | add WordArt to worksheet | Excel watermark programmatically | Aspose.Cells shape ordering
// Common Searches: how to place WordArt behind cells in Aspose.Cells | set shape Z-order in Excel using C# | create background watermark with WordArt in .NET | Aspose.Cells send shape to back | programmatic Excel watermark Aspose
// Developer Intent: Place a WordArt shape behind all worksheet content so it functions as a background watermark.
// Use Cases: Generate confidential reports with a semi‑transparent "CONFIDENTIAL" watermark that does not obscure data. | Add a company logo WordArt as a persistent background element across multiple worksheets. | Automate the insertion of rotated WordArt watermarks into template workbooks while preserving chart and table visibility.
// AI Prompts: Write C# code with Aspose.Cells to add a WordArt shape, set its ZOrderPosition to 0, and save the workbook. | Show how to apply a semi‑transparent WordArt watermark to every sheet in an existing Excel file using Aspose.Cells for .NET. | Explain the ZOrderPosition property in Aspose.Cells and how to use it to send a shape to the back of a worksheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, inserts a WordArt shape with custom text, sets its ZOrderPosition to 0 so it sits behind all other objects, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a WordArt shape (preset style, text, position and size)
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1, // preset style
                "CONFIDENTIAL",                  // text
                0,   // top row index
                0,   // vertical offset (pixels)
                0,   // left column index
                0,   // horizontal offset (pixels)
                200, // height (pixels)
                600  // width (pixels)
            );

            // Send the WordArt to the back so it acts as a background watermark
            // In Aspose.Cells, set ZOrderPosition to a low value to place it behind other shapes
            wordArt.ZOrderPosition = 0;

            // Define output file path
            string outputPath = "WordArtBackground.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
