// Title: Validate that Shape.IsSmartArt returns false for a regular rectangle shape after saving and reloading an XLSX workbook with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to confirm a shape is not SmartArt after workbook save | verify IsSmartArt property for a rectangle shape in a loaded Excel file using Aspose.Cells | determine if a regular shape becomes SmartArt after reloading an XLSX workbook in .NET | Aspose.Cells shape.IsSmartArt false check example for non‑SmartArt shapes | check shape type after round‑trip save and load with Aspose.Cells C#
// Tags: Aspose.Cells shape IsSmartArt verification | C# rectangle shape non‑SmartArt detection | Excel workbook save reload shape property | XLSX shape IsSmartArt false validation | Aspose.Cells shape round‑trip test

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape, saves it as XLSX, reloads the file, retrieves the shape by name, and asserts that the Shape.IsSmartArt property is false, confirming that regular shapes remain non‑SmartArt after a workbook round‑trip.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a regular shape (rectangle) to the worksheet
            // Parameters: upper left row, upper left column, top, left, bottom, right
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 200);
            shape.Name = "RegularRectangle";

            // Save the workbook to a temporary file
            string filePath = "SmartArtCheck.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Ensure the file exists before loading
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            // Load the workbook back
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Retrieve the shape by name
            Shape loadedShape = loadedSheet.Shapes["RegularRectangle"];

            // Validate that IsSmartArt is false for a regular shape
            bool isSmartArt = loadedShape.IsSmartArt;
            Console.WriteLine($"Shape.IsSmartArt = {isSmartArt}");

            if (isSmartArt)
                throw new InvalidOperationException("IsSmartArt should be false for a regular shape.");
            else
                Console.WriteLine("Validation passed: regular shape is not SmartArt.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
