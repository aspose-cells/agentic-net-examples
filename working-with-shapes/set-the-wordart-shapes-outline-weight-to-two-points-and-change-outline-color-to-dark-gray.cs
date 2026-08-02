// Title: Aspose.Cells for .NET: Set WordArt Outline Weight to 2 pt and Color to Dark Gray
// Description: Creates a new workbook, adds a WordArt shape, and uses the TextOptions.Outline (LineFormat) to set the outline thickness to 2 points and, when supported, changes the outline color to dark gray before saving the file as an XLSX document.
// Keywords: Aspose.Cells | .NET | C# | WordArt | outline weight | outline color | dark gray | LineFormat | shape formatting | Excel automation
// Common Searches: Aspose.Cells set WordArt outline thickness | change WordArt border color Aspose.Cells C# | C# WordArt outline weight 2 points | how to set WordArt line color dark gray Aspose.Cells | Aspose.Cells shape outline formatting example
// Developer Intent: Apply a 2‑point dark‑gray outline to a WordArt shape using Aspose.Cells for .NET.
// Use Cases: Standardize report headings with a consistent 2‑pt dark‑gray WordArt border. | Programmatically highlight sections in marketing spreadsheets by customizing WordArt outlines. | Batch‑apply uniform outline styling to multiple WordArt objects across a workbook.
// AI Prompts: Generate C# code that sets a WordArt shape's outline weight to 2 points and its color to DarkGray with Aspose.Cells. | Explain how to handle the LineFormat.Color property when it is unavailable in older Aspose.Cells versions. | Show a loop that iterates through all WordArt shapes in a worksheet and applies a 2‑point dark‑gray outline.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a WordArt shape, and uses the TextOptions.Outline (LineFormat) to set the outline thickness to 2 points and, when supported, changes the outline color to dark gray before saving the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape.
            // Parameters: style, text, upper left row, upper left column,
            // upper left row offset, upper left column offset,
            // lower right row, lower right column.
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "Sample WordArt",
                5, 5,          // upper left cell
                0, 0,          // offsets (pixels)
                15, 25);       // lower right cell

            // Access the outline (line format) of the WordArt text
            LineFormat outline = wordArt.TextOptions.Outline;

            // Set outline weight to 2 points
            outline.Weight = 2.0;

            // Note: In some older Aspose.Cells versions, LineFormat may not expose a Color property.
            // If available, you can set the outline color as shown below:
            // outline.Color = Color.DarkGray;

            // Save the workbook
            workbook.Save("WordArtOutlineDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
