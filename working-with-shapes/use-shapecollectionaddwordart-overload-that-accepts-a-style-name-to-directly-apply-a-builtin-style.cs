// Title: Add WordArt with a Built‑in Preset Style Using ShapeCollection.AddWordArt (Aspose.Cells .NET)
// Description: Demonstrates how to create a workbook, access the first worksheet's ShapeCollection, and call the AddWordArt overload that accepts a PresetWordArtStyle (e.g., WordArtStyle7). The example sets the text, position, size, makes the font bold, adjusts the size to 24 pt, and saves the file as WordArtStyleDemo.xlsx.
// Keywords: Aspose.Cells AddWordArt preset style | ShapeCollection.AddWordArt .NET | WordArtStyle7 Aspose.Cells | apply built‑in WordArt style C# | customize WordArt font Aspose.Cells | Excel WordArt automation | Aspose.Cells shape styling
// Common Searches: ShapeCollection.AddWordArt with PresetWordArtStyle example | how to apply a built‑in WordArt style in Aspose.Cells | C# add WordArt to worksheet using Aspose.Cells | set WordArt font bold size in Aspose.Cells | Aspose.Cells WordArt style parameter
// Developer Intent: Insert a WordArt shape that uses a predefined style and optionally tweak its font attributes programmatically.
// Use Cases: Create a stylized title banner for financial reports. | Add colored WordArt labels to chart legends for visual emphasis. | Programmatically annotate key sections of a spreadsheet with bold, large WordArt text.
// AI Prompts: Generate C# code that adds a WordArt shape with PresetWordArtStyle.WordArtStyle5, sets the text to "Sales Summary", makes the font italic, and uses a 18 pt size. | Write a method that inserts multiple WordArt shapes, each with a different built‑in style, and returns the collection of created Shape objects.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, access the first worksheet's ShapeCollection, and call the AddWordArt overload that accepts a PresetWordArtStyle (e.g., WordArtStyle7). The example sets the text, position, size, makes the font bold, adjusts the size to 24 pt, and saves the file as WordArtStyleDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape using a built‑in preset style.
        // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
            "Aspose WordArt",                 // Text to display
            2,    // Upper‑left row index
            0,    // Vertical offset in pixels
            2,    // Upper‑left column index
            0,    // Horizontal offset in pixels
            100,  // Height in pixels
            400   // Width in pixels
        );

        // Optional: further customize the WordArt appearance
        wordArt.TextEffect.FontBold = true;
        wordArt.TextEffect.FontSize = 24;

        // Save the workbook (lifecycle rule)
        workbook.Save("WordArtStyleDemo.xlsx");
    }
}
