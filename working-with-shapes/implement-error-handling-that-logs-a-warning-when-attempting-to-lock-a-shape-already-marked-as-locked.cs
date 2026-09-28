// Title: Log a warning when attempting to lock an already locked shape with Aspose.Cells for .NET
// AI Prompts: Write C# code that checks a worksheet shape's IsLocked property using Aspose.Cells and writes a warning to the console if the shape is already locked. | Create a reusable C# method that tries to lock an Aspose.Cells shape, logs a warning when the shape is already locked, and handles any exceptions.
// Common Searches: asp.net cells avoid locking a shape that is already locked | c# check shape.IsLocked before setting it in Aspose.Cells | log warning instead of error when shape lock already set Aspose.Cells | prevent duplicate shape lock in Excel workbook using Aspose.Cells C# | exception handling for shape locking in Aspose.Cells
// Tags: Aspose.Cells shape lock verification | C# check shape.IsLocked property | log warning for duplicate shape lock | Excel shape lock handling with Aspose.Cells | exception handling for shape locking

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel file, checks the first shape on the first worksheet, logs a warning if the shape is already locked, otherwise locks it, and saves the workbook.
class ShapeLockHandler
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one shape on the sheet
        if (sheet.Shapes.Count == 0)
        {
            Console.WriteLine("No shapes found on the worksheet.");
            return;
        }

        // Get the first shape
        Shape shape = sheet.Shapes[0];

        try
        {
            // Lock the shape if not already locked
            if (shape.IsLocked)
            {
                Console.WriteLine($"Warning: Shape '{shape.Name}' is already locked.");
            }
            else
            {
                shape.IsLocked = true;
                Console.WriteLine($"Shape '{shape.Name}' has been locked successfully.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while processing shape '{shape.Name}': {ex.Message}");
        }

        try
        {
            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
