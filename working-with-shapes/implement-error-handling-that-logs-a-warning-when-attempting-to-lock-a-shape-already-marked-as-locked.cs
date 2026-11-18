// Title: C# Aspose.Cells example: log a warning if a worksheet shape is already locked before setting IsLocked
// Common Searches: how to check if a shape is locked before setting IsLocked in Aspose.Cells C# | Aspose.Cells C# log warning when trying to lock an already locked worksheet shape | prevent double locking of Excel shapes using Aspose.Cells API | C# console output warning for locked shape in an Aspose.Cells workbook | example of using shape.IsLocked property with Aspose.Cells in C#
// Tags: Aspose.Cells shape lock validation | C# Aspose.Cells shape IsLocked check | logging warnings for Aspose.Cells shape operations | Excel workbook shape locking with Aspose.Cells | prevent duplicate shape lock Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code creates a placeholder workbook if needed, loads an existing Excel file, retrieves the first shape on the first worksheet, checks the shape's IsLocked property, writes a console warning when the shape is already locked, otherwise locks the shape, saves the workbook, and handles any exceptions.
class ShapeLockExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input workbook exists; create a simple one with a shape if missing
            if (!File.Exists(inputPath))
            {
                var newWb = new Workbook();
                var ws = newWb.Worksheets[0];
                // Add a rectangle shape as a placeholder
                var placeholder = ws.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                placeholder.Name = "PlaceholderShape";
                newWb.Save(inputPath);
                Console.WriteLine($"Created placeholder workbook '{inputPath}'.");
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Get the first shape
            Shape shape = sheet.Shapes[0];
            int shapeIndex = sheet.Shapes.IndexOf(shape); // Retrieve the shape's index safely

            // Check if the shape is already locked
            if (shape.IsLocked)
            {
                Console.WriteLine($"Warning: Shape '{shape.Name}' (Index {shapeIndex}) is already locked.");
            }
            else
            {
                // Lock the shape
                shape.IsLocked = true;
                Console.WriteLine($"Shape '{shape.Name}' (Index {shapeIndex}) has been locked.");
            }

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
