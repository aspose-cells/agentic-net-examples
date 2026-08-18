// Title: Aspose.Cells for .NET: Set WordArt Shape Outline Weight to 2 pt and Color to Dark Gray
// Description: Creates a new workbook, adds a WordArt shape, and uses the Shape.LineFormat property to set a 2‑point outline thickness and dark‑gray border before saving the file as an XLSX workbook.
// Keywords: Aspose.Cells | C# | .NET | WordArt shape | outline weight | line format | dark gray border | Excel shape styling | Shape.LineFormat | border thickness
// Common Searches: Aspose.Cells set WordArt outline thickness .NET | Change WordArt border color to dark gray using Aspose.Cells | How to adjust WordArt shape line weight in C# | Aspose.Cells LineFormat example for WordArt | Set shape outline properties in Aspose.Cells workbook
// Developer Intent: Apply a 2‑point outline and dark‑gray color to a WordArt shape in an Excel file using Aspose.Cells for .NET.
// Use Cases: Design report titles with a subtle dark‑gray border for brand consistency. | Automate generation of Excel dashboards where WordArt headings must follow a specific outline style. | Batch‑apply uniform outline formatting to multiple WordArt objects across worksheets.
// AI Prompts: Show C# code that sets a WordArt shape's LineFormat.Weight to 2 points and LineFormat.ForeColor to DarkGray with Aspose.Cells. | Provide a script to loop through all WordArt shapes in a worksheet and apply a 2‑point dark‑gray outline. | Explain the relationship between LineFormat.Weight, LineFormat.ForeColor, and other line‑format properties in Aspose.Cells.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a WordArt shape, and uses the Shape.LineFormat property to set a 2‑point outline thickness and dark‑gray border before saving the file as an XLSX workbook.
class WordArtOutlineExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Sample WordArt",                 // text
            1, 0,                             // top row, top offset (pixels)
            1, 0,                             // left column, left offset (pixels)
            100, 400);                        // height, width (pixels)

        // Set the outline (border) weight to 2 points
        wordArt.LineFormat.Weight = 2; // weight in points

        // Change the outline color to dark gray
        wordArt.LineFormat.ForeColor = Color.DarkGray;

        // Save the workbook
        workbook.Save("WordArtOutlineDemo.xlsx");
    }
}
