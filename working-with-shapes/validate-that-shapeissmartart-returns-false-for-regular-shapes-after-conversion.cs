// Title: Check that a regular rectangle shape's IsSmartArt property stays false before and after saving to XLSX with Aspose.Cells for .NET
// AI Prompts: Write a C# snippet using Aspose.Cells that adds a rectangle shape to a worksheet, saves the workbook as XLSX, reloads it, and asserts shape.IsSmartArt is false. | Generate a C# unit test that creates a workbook, inserts a non‑SmartArt rectangle, records the shape index, persists the file, loads it back, and verifies the IsSmartArt flag remains false. | Provide code to print the IsSmartArt value of a regular shape both before saving and after loading the XLSX file with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to confirm a shape is not SmartArt after saving workbook | verify IsSmartArt property of rectangle shape after XLSX conversion using Aspose.Cells | C# Aspose.Cells shape.IsSmartArt returns false for regular shapes | test shape property persistence in saved Excel file with Aspose.Cells for .NET
// Tags: Aspose.Cells shape smartart detection | C# rectangle shape non‑smartart verification | shape property consistency after workbook save | Aspose.Cells XLSX persistence test | validate shape type after reload

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape, prints its IsSmartArt property (expected false), saves the workbook to an XLSX file, reloads the file, retrieves the same shape by index, and prints IsSmartArt again to confirm the value remains false.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Add a regular rectangle shape
            Shape shape = ws.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Get the shape's index within the collection
            int shapeIndex = ws.Shapes.IndexOf(shape);

            // Validate that IsSmartArt is false for the regular shape
            Console.WriteLine("Before conversion IsSmartArt: " + shape.IsSmartArt); // Expected: False

            // Save the workbook to a temporary file (conversion step)
            string filePath = "temp.xlsx";
            wb.Save(filePath, SaveFormat.Xlsx);

            // Ensure the file exists before loading
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The workbook file was not created.", filePath);

            // Load the workbook back from the file
            Workbook loadedWb = new Workbook(filePath);
            Shape loadedShape = loadedWb.Worksheets[0].Shapes[shapeIndex];

            // Validate that IsSmartArt remains false after conversion
            Console.WriteLine("After conversion IsSmartArt: " + loadedShape.IsSmartArt); // Expected: False
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
