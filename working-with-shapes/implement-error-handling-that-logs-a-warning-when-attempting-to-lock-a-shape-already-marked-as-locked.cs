// Title: Log a warning when attempting to lock an already locked shape using Aspose.Cells for .NET
// AI Prompts: Write C# code that checks a worksheet shape's IsLocked property and logs a warning instead of re‑locking it with Aspose.Cells. | Create a reusable C# method that locks an Aspose.Cells shape only if it is not already locked and writes a warning when the shape is already locked. | Show how to handle duplicate shape‑lock attempts in Aspose.Cells by using conditional logic and console logging.
// Common Searches: Aspose.Cells C# how to avoid locking a shape that is already locked | C# log warning when shape.IsLocked is true in Aspose.Cells workbook | check shape lock status before setting IsLocked property Aspose.Cells .NET | prevent duplicate shape locking Aspose.Cells example code | Aspose.Cells shape lock validation and warning message
// Tags: Aspose.Cells shape lock validation | C# check shape.IsLocked before setting | Aspose.Cells warning log for duplicate lock | Excel shape lock handling with Aspose.Cells | conditional shape locking Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, verifies that a shape exists on the first worksheet, checks the shape's IsLocked property, logs a warning if the shape is already locked, otherwise locks it, and then saves the workbook, with basic file existence checks and exception handling.
class ShapeLockHandler
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Get the first shape (you can select any shape by index)
            Shape shape = sheet.Shapes[0];

            // Check if the shape is already locked
            if (shape.IsLocked)
            {
                // Log a warning instead of attempting to lock again
                Console.WriteLine($"Warning: Shape '{shape.Name}' is already locked.");
            }
            else
            {
                // Lock the shape
                shape.IsLocked = true;
                Console.WriteLine($"Shape '{shape.Name}' has been locked.");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
