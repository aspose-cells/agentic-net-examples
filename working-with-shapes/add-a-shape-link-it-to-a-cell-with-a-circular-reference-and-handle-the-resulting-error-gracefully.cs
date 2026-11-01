// Title: Add a rectangle shape to a worksheet, create a circular reference in a cell, and handle the calculation error with Aspose.Cells for .NET (C#)
// Common Searches: how to add a rectangle shape to an Aspose.Cells worksheet in C# | handling circular reference errors during workbook.CalculateFormula in Aspose.Cells | saving an Aspose.Cells workbook to a dynamic folder path in C# | positioning a shape without SetLinkToCell using Aspose.Cells C# | catching formula calculation exceptions for circular references in Aspose.Cells
// Tags: add drawing shape Aspose.Cells C# | circular reference handling Aspose.Cells | formula calculation exception handling Aspose.Cells | shape positioning alternative method Aspose.Cells | save workbook with folder creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, inserts a rectangle drawing shape onto the first worksheet, introduces a circular reference in cell A1, attempts to calculate formulas inside a nested try‑catch to capture the generic exception thrown for circular references, ensures the output directory exists, and saves the workbook while handling any I/O errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

            // NOTE: The SetLinkToCell method is not available in the current Aspose.Cells version.
            // If linking a shape to a cell is required, use alternative positioning methods.
            // shape.SetLinkToCell("A1", 0, 0); // Removed to avoid compilation error

            // Introduce a circular reference in cell A1 (e.g., =A1+1)
            Cell targetCell = sheet.Cells["A1"];
            targetCell.Formula = "=A1+1";

            // Attempt to calculate formulas and handle circular reference errors gracefully
            try
            {
                workbook.CalculateFormula();
                Console.WriteLine("Formula calculation completed successfully.");
            }
            catch (Exception ex)
            {
                // Aspose.Cells throws a generic exception for circular references.
                // Handle it by informing the user without crashing the application.
                Console.WriteLine("An error occurred during formula calculation:");
                Console.WriteLine(ex.Message);
            }

            // Define output file path
            string outputPath = "ShapeWithCircularReference.xlsx";

            // Ensure the directory exists before saving (handle case where outputPath has no directory)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file with error handling
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine("Unexpected error:");
            Console.WriteLine(ex.Message);
        }
    }
}
