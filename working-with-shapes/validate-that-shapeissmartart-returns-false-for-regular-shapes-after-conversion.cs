// Title: Validate Shape.IsSmartArt is false for regular shapes and GetResultOfSmartArt returns null – Aspose.Cells for .NET
// Description: Creates a workbook, adds a standard rectangle shape, confirms the IsSmartArt property is false, calls GetResultOfSmartArt (which should return null for non‑SmartArt shapes), prints the outcomes, and saves the file.
// Keywords: Aspose.Cells | Shape.IsSmartArt | GetResultOfSmartArt | C# example | non‑SmartArt shape | Excel shape validation | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells check if shape is SmartArt | GetResultOfSmartArt null for rectangle | Shape.IsSmartArt false example | C# Aspose.Cells shape type verification | How to detect SmartArt in Excel with Aspose
// Developer Intent: Verify that a regular rectangle is identified as non‑SmartArt and that attempting to retrieve SmartArt data returns null, preventing runtime errors.
// Use Cases: Pre‑validation before performing SmartArt‑specific operations. | Unit testing shape classification in automated Excel processing pipelines. | Debugging shape handling logic when migrating spreadsheets to Aspose.Cells.
// AI Prompts: Generate an MSTest method that asserts Shape.IsSmartArt is false and GetResultOfSmartArt returns null for a newly added rectangle using Aspose.Cells. | Write a C# loop that scans all worksheet shapes, logs each shape's IsSmartArt status, and safely skips null results from GetResultOfSmartArt. | Explain why GetResultOfSmartArt returns null for non‑SmartArt shapes and show best‑practice code to check this condition before casting.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a standard rectangle shape, confirms the IsSmartArt property is false, calls GetResultOfSmartArt (which should return null for non‑SmartArt shapes), prints the outcomes, and saves the file.
class ValidateSmartArt
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a regular rectangle shape (non‑SmartArt)
        Shape rectangle = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

        // Check that IsSmartArt is false for the regular shape
        bool isSmartArt = rectangle.IsSmartArt;
        Console.WriteLine($"Rectangle IsSmartArt: {isSmartArt}");

        // Attempt to convert the shape to a GroupShape (should return null for non‑SmartArt)
        GroupShape group = rectangle.GetResultOfSmartArt();
        Console.WriteLine($"GetResultOfSmartArt returned null: {group == null}");

        // Save the workbook (optional verification step)
        workbook.Save("ValidateSmartArt.xlsx");
    }
}
