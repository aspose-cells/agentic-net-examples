// Title: Add WordArt with a Built‑In Preset Style Using ShapeCollection.AddWordArt in Aspose.Cells for .NET (C#)
// Description: This example shows how to create a new Workbook, access the first worksheet's ShapeCollection, and call the AddWordArt overload that accepts a PresetWordArtStyle. The code inserts a WordArt shape with the built‑in Gold fill (WordArtStyle5), optionally sets its rotation, and saves the file as WordArtWithPresetStyle.xlsx.
// Keywords: Aspose.Cells | C# | AddWordArt | ShapeCollection | PresetWordArtStyle | WordArtStyle5 | built‑in WordArt style | Excel shape example | Aspose.Cells for .NET tutorial | GitHub Aspose.Cells sample | download Aspose.Cells code
// Common Searches: Aspose.Cells add WordArt with preset style C# | ShapeCollection.AddWordArt PresetWordArtStyle example | How to apply built‑in WordArt style in Aspose.Cells | Set rotation for WordArt shape using Aspose.Cells | C# code to insert WordArt into Excel workbook
// Developer Intent: Insert a WordArt shape and apply a built‑in preset style in a single call.
// Use Cases: Create a decorative title banner for a financial report using the gold WordArtStyle5. | Automate the addition of styled WordArt headings across multiple worksheets. | Highlight key metrics with rotated WordArt labels positioned over specific cells.
// AI Prompts: Generate C# code that adds WordArt with PresetWordArtStyle.WordArtStyle3 and rotates it 45 degrees using Aspose.Cells. | Provide a loop that reads a list of strings and adds each as a WordArt shape with a different preset style to a worksheet. | Explain how to enumerate all PresetWordArtStyle values and select one based on user input in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example shows how to create a new Workbook, access the first worksheet's ShapeCollection, and call the AddWordArt overload that accepts a PresetWordArtStyle. The code inserts a WordArt shape with the built‑in Gold fill (WordArtStyle5), optionally sets its rotation, and saves the file as WordArtWithPresetStyle.xlsx.
class AddWordArtExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape using a built‑in preset style
        // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle5, // Gold fill style
            "Aspose WordArt",                 // Text to display
            2,    // Upper‑left row index
            0,    // Vertical offset in pixels
            2,    // Upper‑left column index
            0,    // Horizontal offset in pixels
            100,  // Height in pixels
            400   // Width in pixels
        );

        // Example: set rotation if desired
        wordArt.RotationAngle = 0;

        // Save the workbook with the WordArt shape
        workbook.Save("WordArtWithPresetStyle.xlsx");
    }
}
