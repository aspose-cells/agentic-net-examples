// Title: Shape.IsSmartArt remains false for a regular rectangle after GetResultOfSmartArt in Aspose.Cells for .NET
// Description: Creates a workbook, adds a standard rectangle shape, checks that Shape.IsSmartArt is false, calls GetResultOfSmartArt (which returns null for non‑SmartArt shapes), and confirms the property stays false before saving the file.
// Keywords: Aspose.Cells | .NET | C# | Shape.IsSmartArt | SmartArt detection | GetResultOfSmartArt | regular shape | rectangle shape | non‑SmartArt conversion | shape type verification
// Common Searches: Shape.IsSmartArt false example Aspose.Cells | GetResultOfSmartArt returns null | how to check if shape is SmartArt in C# | Aspose.Cells verify shape type | C# SmartArt conversion test
// Developer Intent: Confirm that calling GetResultOfSmartArt does not alter the IsSmartArt flag for a non‑SmartArt shape.
// Use Cases: Prevent accidental SmartArt processing on standard shapes during workbook automation. | Safely attempt SmartArt conversion on any shape without affecting regular shapes. | Log or filter shapes based on SmartArt status to avoid runtime errors.
// AI Prompts: Generate an NUnit test that adds a rectangle, asserts Shape.IsSmartArt is false before and after GetResultOfSmartArt, and verifies the method returns null. | Write C# code that iterates all worksheet shapes, prints each shape's IsSmartArt value, and handles null results from GetResultOfSmartArt. | Create a helper method that checks IsSmartArt, performs SmartArt conversion only when true, and returns the original shape unchanged for non‑SmartArt objects.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtCheck
{
    // Creates a workbook, adds a standard rectangle shape, checks that Shape.IsSmartArt is false, calls GetResultOfSmartArt (which returns null for non‑SmartArt shapes), and confirms the property stays false before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a regular rectangle shape (non‑SmartArt)
            Shape rectangle = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

            // Verify that the shape is not a SmartArt before any conversion
            bool isSmartBefore = rectangle.IsSmartArt; // Expected: false
            Console.WriteLine($"Rectangle IsSmartArt before conversion: {isSmartBefore}");

            // Attempt to convert the shape using GetResultOfSmartArt (returns null for non‑SmartArt)
            GroupShape resultGroup = rectangle.GetResultOfSmartArt();
            Console.WriteLine($"GetResultOfSmartArt returned: {(resultGroup == null ? "null" : "GroupShape")}");

            // Verify that IsSmartArt is still false after the conversion attempt
            bool isSmartAfter = rectangle.IsSmartArt; // Expected: false
            Console.WriteLine($"Rectangle IsSmartArt after conversion: {isSmartAfter}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("RegularShapeCheck.xlsx");
        }
    }
}
