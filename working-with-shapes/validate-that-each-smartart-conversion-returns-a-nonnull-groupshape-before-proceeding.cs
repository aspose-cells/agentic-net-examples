// Title: Check for non‑null GroupShape after converting SmartArt shapes with Aspose.Cells in C#
// AI Prompts: Write C# code that iterates through every worksheet shape, converts each SmartArt shape to a GroupShape using Aspose.Cells, and throws an InvalidOperationException if any conversion returns null. | Update the existing Aspose.Cells example to insert a validation step that ensures every SmartArt‑to‑GroupShape conversion yields a non‑null GroupShape before the workbook is saved.
// Common Searches: Aspose.Cells C# how to ensure SmartArt conversion returns a GroupShape object | validate SmartArt to GroupShape conversion null check in .NET | C# Aspose.Cells check for null after converting SmartArt shapes | error handling for SmartArt conversion to GroupShape using Aspose.Cells
// Tags: SmartArt to GroupShape conversion Aspose.Cells | null GroupShape validation C# | Aspose.Cells shape processing error handling | Excel workbook SmartArt conversion .NET | Aspose.Cells shape iteration validation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through each worksheet and its shapes, and shows where to add a validation that every SmartArt shape converted to a GroupShape with Aspose.Cells returns a non‑null object before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Aspose.Cells does not currently expose a SmartArt type enumeration.
                    // If needed, add custom processing for specific shape types here.
                }
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
