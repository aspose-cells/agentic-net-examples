// Title: Check that a regular rectangle shape remains non‑SmartArt after saving and loading an XLSX workbook with Aspose.Cells for .NET
// Common Searches: how to determine if a shape is SmartArt using Aspose.Cells in C# | verify Shape.IsSmartArt returns false for a rectangle after workbook save and load | retrieve a shape by its name from a loaded XLSX file with Aspose.Cells | test whether shape type persists after converting a workbook to XLSX in C# | C# example for checking SmartArt flag on regular shapes in an Excel workbook
// Tags: Aspose.Cells Shape.IsSmartArt verification | C# rectangle shape non‑SmartArt check | workbook save and load shape property persistence | XLSX conversion shape type validation | Aspose.Cells retrieve shape by name

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, adds a rectangle shape named "MyRectangle" to the first worksheet, saves the workbook to a memory stream as XLSX, reloads it, retrieves the shape by its name, and confirms that the Shape.IsSmartArt property is false, demonstrating that regular shapes are not mistakenly identified as SmartArt after conversion.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a regular rectangle shape to the worksheet
            // AddShape returns a Shape object in recent versions
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // top offset (in pixels)
                0,   // left offset (in pixels)
                100, // height (in pixels)
                50   // width (in pixels)
            );

            // Assign a unique name to retrieve the shape after loading
            const string shapeName = "MyRectangle";
            shape.Name = shapeName;

            // Save the workbook to a memory stream (save rule)
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsx);
                stream.Position = 0;

                // Load the workbook from the memory stream (load rule)
                Workbook loadedWorkbook = new Workbook(stream);
                Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

                // Retrieve the shape by its name
                Shape loadedShape = loadedWorksheet.Shapes[shapeName];

                // Validate that IsSmartArt is false for the regular shape after conversion
                bool isSmartArt = loadedShape.IsSmartArt;
                Console.WriteLine($"IsSmartArt: {isSmartArt}");
                if (!isSmartArt)
                {
                    Console.WriteLine("Validation passed: Regular shape is not SmartArt.");
                }
                else
                {
                    Console.WriteLine("Validation failed: Shape incorrectly identified as SmartArt.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
