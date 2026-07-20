// Title: Add a WordArt Watermark to the First Worksheet in a New Excel Workbook (Aspose.Cells for .NET)
// Description: Creates a new Workbook, inserts a WordArt shape with the text "CONFIDENTIAL" as a background watermark on the first worksheet, sets its Z‑order behind cells, and saves the file as WordArtWatermark.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | WordArt watermark | Excel shape | ZOrderPosition | background watermark | add WordArt to worksheet | .NET Excel API | create workbook | shape collection
// Common Searches: Aspose.Cells add WordArt watermark C# | how to place a WordArt shape behind cells in Excel | set ZOrderPosition for Excel shape Aspose.Cells | create Excel file with confidential watermark using .NET | insert WordArt as background in first worksheet
// Developer Intent: Generate an Excel workbook and place a WordArt shape as a background watermark on the first sheet.
// Use Cases: Produce confidential reports with a "CONFIDENTIAL" watermark on the initial sheet. | Design branded templates that display a company slogan as a WordArt overlay. | Automate batch creation of Excel files that include a legal disclaimer watermark.
// AI Prompts: Show code to add a semi‑transparent WordArt watermark to every worksheet in an existing workbook with Aspose.Cells for .NET. | Provide an example that rotates a WordArt watermark and changes its color using Aspose.Cells. | Explain how to position a WordArt watermark relative to page margins instead of cell coordinates in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, inserts a WordArt shape with the text "CONFIDENTIAL" as a background watermark on the first worksheet, sets its Z‑order behind cells, and saves the file as WordArtWatermark.xlsx using Aspose.Cells for .NET.
class WordArtWatermarkDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add a WordArt shape as a watermark
            // Parameters: style, text, topRow, top (pixels), leftColumn, left (pixels), height (pixels), width (pixels)
            Shape wordArt = shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle2, // preset style
                "CONFIDENTIAL",                  // watermark text
                5, 0,                            // top row index and vertical offset
                5, 0,                            // left column index and horizontal offset
                100, 400                         // height and width of the shape
            );

            // Send the shape to the background so it appears as a watermark
            // Lower ZOrderPosition values are rendered behind higher values
            wordArt.ZOrderPosition = 0;

            // Save the workbook to a file
            workbook.Save("WordArtWatermark.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
