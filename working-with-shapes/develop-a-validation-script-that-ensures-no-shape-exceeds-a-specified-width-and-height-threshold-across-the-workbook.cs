// Title: C# script to validate that every shape in an Excel workbook stays within a maximum width and height using Aspose.Cells
// AI Prompts: Write a C# console program that loads an Excel workbook with Aspose.Cells, iterates all worksheets and their Shapes collections, and prints any shape whose Width or Height exceeds specified point limits. | Adapt the shape‑validation tool to accept custom max width and max height values via command‑line arguments and output the validation results to a log file. | Enhance the validator to skip hidden shapes (IsVisible = false) while still checking the dimensions of visible shapes.
// Common Searches: aspocells c# check shape size limits in workbook | how to iterate over all shapes in an Excel file using Aspose.Cells | C# program to find oversized shapes in .xlsx with Aspose.Cells | validate shape dimensions against max width height in Aspose.Cells .NET | log shape dimension violations in Excel using Aspose.Cells API
// Tags: Aspose.Cells iterate worksheet shapes | Aspose.Cells shape dimension validation | C# validate Excel shape size limits | Aspose.Cells shape size check .xlsx | console shape validation script .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// A C# console application that receives an Excel file path, loads the workbook with Aspose.Cells, walks through each worksheet and its Shapes collection, and reports any shape whose Width exceeds 500 points or Height exceeds 300 points. If no violations are found, it confirms all shapes meet the defined size constraints.
class ShapeSizeValidator
{
    // Maximum allowed dimensions (in points)
    const double MaxWidth = 500.0;
    const double MaxHeight = 300.0;

    static void Main(string[] args)
    {
        // Expect the workbook file path as the first argument
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: ShapeSizeValidator <workbook_path>");
            return;
        }

        string workbookPath = args[0];

        // Verify that the file exists before attempting to load
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: File not found – \"{workbookPath}\"");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        bool anyViolation = false;

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Check width and height against the thresholds
                    if (shape.Width > MaxWidth || shape.Height > MaxHeight)
                    {
                        anyViolation = true;
                        Console.WriteLine(
                            $"Violation in sheet \"{sheet.Name}\": Shape \"{shape.Name}\" " +
                            $"has size {shape.Width}x{shape.Height} points, " +
                            $"exceeds allowed {MaxWidth}x{MaxHeight} points.");
                    }
                }
            }

            if (!anyViolation)
            {
                Console.WriteLine("All shapes are within the specified size limits.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing shapes: {ex.Message}");
        }
    }
}
