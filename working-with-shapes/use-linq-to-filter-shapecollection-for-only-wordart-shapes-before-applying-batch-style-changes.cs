// Title: LINQ Filter WordArt Shapes and Apply Batch Text Effects with Aspose.Cells (.NET)
// Description: Creates a workbook, adds WordArt and regular shapes, uses LINQ to select only WordArt objects from the worksheet's ShapeCollection, updates their TextEffectFormat (bold, italic, 24 pt Arial), and saves the result as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | LINQ ShapeCollection | WordArt filter | batch text effect | IsWordArt | TextEffectFormat | Excel shape styling
// Common Searches: filter WordArt shapes Aspose.Cells C# | apply batch formatting to WordArt with LINQ | select only WordArt objects in worksheet | change font of all WordArt shapes programmatically | Aspose.Cells ShapeCollection LINQ example
// Developer Intent: Select only WordArt shapes in a worksheet and modify their text‑effect properties in one pass.
// Use Cases: Standardize WordArt typography across a corporate template. | Increase readability by enlarging and bolding all WordArt before PDF export. | Apply brand‑consistent font styling to WordArt while preserving other drawings.
// AI Prompts: Generate C# code that uses Aspose.Cells to filter WordArt shapes with LINQ and set their TextEffectFormat to bold, italic, 24 pt Arial. | Show an example of enumerating only WordArt objects in a worksheet and updating their font properties in bulk. | Demonstrate how to cast ShapeCollection to Shape, use IsWordArt, and apply batch style changes to WordArt in Aspose.Cells for .NET.

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds WordArt and regular shapes, uses LINQ to select only WordArt objects from the worksheet's ShapeCollection, updates their TextEffectFormat (bold, italic, 24 pt Arial), and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add sample shapes: two WordArt shapes and one regular rectangle
        shapes.AddWordArt(PresetWordArtStyle.WordArtStyle1, "Hello", 0, 0, 0, 0, 100, 200);
        shapes.AddRectangle(2, 0, 2, 0, 100, 100);
        shapes.AddWordArt(PresetWordArtStyle.WordArtStyle5, "World", 5, 0, 5, 0, 100, 200);

        // Use LINQ to filter only the WordArt shapes
        var wordArtShapes = shapes.Cast<Shape>().Where(s => s.IsWordArt);

        // Apply batch style changes to each WordArt shape
        foreach (var shape in wordArtShapes)
        {
            TextEffectFormat textEffect = shape.TextEffect;
            textEffect.FontBold = true;
            textEffect.FontItalic = true;
            textEffect.FontSize = 24;
            textEffect.FontName = "Arial";
        }

        // Save the workbook with the updated shapes
        workbook.Save("WordArtBatchStyle.xlsx");
    }
}
