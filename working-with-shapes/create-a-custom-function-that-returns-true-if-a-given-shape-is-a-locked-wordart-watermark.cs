// Title: C# Extension Method IsLockedWordArtWatermark for Aspose.Cells – Detect Locked WordArt Watermarks
// Description: Defines a ShapeExtensions class with the IsLockedWordArtWatermark extension that returns true only when a Shape is a WordArt object and its IsLocked flag is set. The sample creates a workbook, adds a WordArt watermark, locks it, adds a regular locked shape, and prints the detection results.
// Keywords: Aspose.Cells | C# | Shape extension | WordArt | watermark detection | IsLocked | IsWordArt | Excel shape | locked watermark | extension method
// Common Searches: Aspose.Cells check if shape is locked WordArt | C# detect WordArt watermark in Excel using Aspose | IsLockedWordArtWatermark example | How to identify locked WordArt shapes with Aspose.Cells | extension method for shape watermark detection Aspose
// Developer Intent: Determine whether a given Shape instance represents a locked WordArt watermark.
// Use Cases: Validate that only WordArt watermarks are locked before protecting a worksheet. | Iterate through all worksheet shapes and apply formatting only to unlocked, non‑watermark objects. | Generate an audit report of every locked WordArt watermark across a workbook.
// AI Prompts: Write a C# extension method for Aspose.Cells that returns true if a Shape is a locked WordArt watermark. | Show code that loops through all shapes in a worksheet and lists those satisfying IsLockedWordArtWatermark. | Explain how to protect an Excel sheet while keeping locked WordArt watermarks uneditable using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Defines a ShapeExtensions class with the IsLockedWordArtWatermark extension that returns true only when a Shape is a WordArt object and its IsLocked flag is set. The sample creates a workbook, adds a WordArt watermark, locks it, adds a regular locked shape, and prints the detection results.
public static class ShapeExtensions
{
    // Returns true if the shape is a WordArt and is locked (cannot be modified when the sheet is protected)
    public static bool IsLockedWordArtWatermark(this Shape shape)
    {
        if (shape == null) return false;
        return shape.IsWordArt && shape.IsLocked;
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape (commonly used as a watermark)
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,
            "Sample Watermark",
            0, 0, 200, 100, 0, 0);

        // Lock the WordArt shape
        wordArt.IsLocked = true;

        // Add a regular rectangle shape for comparison
        Shape rectangle = worksheet.Shapes.AddRectangle(1, 0, 150, 100, 100, 100);
        rectangle.IsLocked = true;

        // Use the custom function to check if each shape is a locked WordArt watermark
        Console.WriteLine("WordArt shape is a locked watermark: " + wordArt.IsLockedWordArtWatermark());
        Console.WriteLine("Rectangle shape is a locked watermark: " + rectangle.IsLockedWordArtWatermark());
    }
}
