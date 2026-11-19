// Title: C# script to validate that no shape exceeds a specified width or height in an Excel workbook using Aspose.Cells
// Common Searches: C# Aspose.Cells check shape dimensions across all worksheets | how to enforce maximum shape width and height in an Excel file with Aspose.Cells | validate shape size limits in .xlsx using Aspose.Cells for .NET | Aspose.Cells iterate shapes and report oversized objects in a workbook | C# code to list shapes larger than given dimensions in Excel
// Tags: shape dimension validation Aspose.Cells | iterate worksheet shapes C# | max shape width height points | excel workbook shape size check | collect shape size errors Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, defines maximum width and height in points, iterates through every worksheet and each shape, records any shape whose Width or Height exceeds the limits, and outputs either a success message or a detailed list of oversized shapes while handling missing files and runtime exceptions.
class ShapeValidator
{
    static void Main(string[] args)
    {
        // Path to the Excel file to validate
        string filePath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file '{filePath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Define maximum allowed dimensions (in points)
            double maxWidth = 200.0;   // Example width threshold
            double maxHeight = 150.0;  // Example height threshold

            // Collect validation messages
            List<string> errors = new List<string>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Check if the shape exceeds the allowed width or height
                    if (shape.Width > maxWidth || shape.Height > maxHeight)
                    {
                        string message = $"Worksheet '{sheet.Name}', Shape '{shape.Name}' exceeds size limits. " +
                                         $"Width: {shape.Width}, Height: {shape.Height}";
                        errors.Add(message);
                    }
                }
            }

            // Output validation results
            if (errors.Count == 0)
            {
                Console.WriteLine("All shapes are within the specified size limits.");
            }
            else
            {
                Console.WriteLine("Shape size validation errors:");
                foreach (string err in errors)
                {
                    Console.WriteLine(err);
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
