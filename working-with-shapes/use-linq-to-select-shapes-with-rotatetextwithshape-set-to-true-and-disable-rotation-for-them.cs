// Title: C# LINQ to disable RotateTextWithShape for shapes in Aspose.Cells
// Description: Loads an Excel file, uses LINQ to find all shapes on the first worksheet whose TextBody.TextAlignment.RotateTextWithShape property is true, sets the property to false, and saves the workbook.
// Keywords: Aspose.Cells | C# | LINQ | RotateTextWithShape | shape text alignment | disable text rotation | Excel shape properties | filter shapes
// Common Searches: Aspose.Cells disable RotateTextWithShape | LINQ query for shapes with RotateTextWithShape true | C# set RotateTextWithShape false | how to turn off text rotation in Excel shapes using Aspose | batch update shape text alignment Aspose.Cells
// Developer Intent: Turn off the RotateTextWithShape flag for every shape that currently has it enabled.
// Use Cases: Ensure shape text stays horizontal after rotating the shape. | Standardize text orientation across a workbook template. | Automate cleanup of legacy spreadsheets where text rotation was unintentionally applied.
// AI Prompts: Write C# code with Aspose.Cells that uses LINQ to set RotateTextWithShape = false for matching shapes. | Show a LINQ expression to filter shapes where TextBody.TextAlignment.RotateTextWithShape is true. | Explain null‑checking strategies for TextBody and TextAlignment when updating shape properties in Aspose.Cells.

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Loads an Excel file, uses LINQ to find all shapes on the first worksheet whose TextBody.TextAlignment.RotateTextWithShape property is true, sets the property to false, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Find all shapes where RotateTextWithShape is true and disable it
        var shapesToUpdate = worksheet.Shapes
            .Cast<Shape>()
            .Where(s => s.TextBody != null &&
                        s.TextBody.TextAlignment != null &&
                        s.TextBody.TextAlignment.RotateTextWithShape);

        foreach (var shape in shapesToUpdate)
        {
            shape.TextBody.TextAlignment.RotateTextWithShape = false;
        }

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
