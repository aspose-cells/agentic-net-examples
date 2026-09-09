// Title: C# script using Aspose.Cells to validate that all shapes in an Excel workbook stay within specified width and height limits
// AI Prompts: Write a C# console program with Aspose.Cells that opens a given .xlsx file, iterates through every worksheet, and prints the names and dimensions of shapes whose Width or Height exceed user‑defined thresholds. | Enhance the shape‑validation code to export the list of oversized shapes to a CSV file, including worksheet name, shape name, width, and height. | Add command‑line arguments to the validator so that the workbook path, maximum width, and maximum height can be supplied at runtime, and return a non‑zero exit code when any violations are detected.
// Common Searches: c# check if any shape in an Excel workbook exceeds max width or height using Aspose.Cells | how to list oversized shapes in a .xlsx file with Aspose.Cells .NET | asp.net cells shape size validation script for Excel files | automate enforcement of shape dimension limits in Excel via C# Aspose.Cells
// Tags: Aspose.Cells shape dimension validation | C# iterate worksheet shapes | Excel shape size limit enforcement | Aspose.Cells shape width height check | C# console shape validator for .xlsx

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook with Aspose.Cells, walks through each worksheet's Shapes collection, and reports any shape whose Width or Height exceeds the defined maximum values (e.g., 200 pt width, 150 pt height). Violations are output to the console, and the script can be extended to log results or accept configurable thresholds.
class ShapeValidator
{
    static void Main()
    {
        // Path to the workbook to validate
        string workbookPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
            return;
        }

        // Maximum allowed dimensions (in points)
        double maxWidth = 200.0;
        double maxHeight = 150.0;

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Collect any shape violations
            List<string> violations = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Check if the shape exceeds width or height limits
                    if (shape.Width > maxWidth || shape.Height > maxHeight)
                    {
                        string message = $"Sheet '{sheet.Name}', Shape '{shape.Name}' exceeds limits: Width={shape.Width}, Height={shape.Height}";
                        violations.Add(message);
                    }
                }
            }

            // Output validation results
            if (violations.Count == 0)
            {
                Console.WriteLine("No shapes exceed the specified width and height thresholds.");
            }
            else
            {
                Console.WriteLine("Shapes exceeding the specified dimensions:");
                foreach (string v in violations)
                {
                    Console.WriteLine(v);
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
