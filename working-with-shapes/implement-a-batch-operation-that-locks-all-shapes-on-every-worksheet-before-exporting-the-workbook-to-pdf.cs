// Title: Lock all shapes on every worksheet in an Excel workbook with Aspose.Cells for .NET before saving as PDF
// AI Prompts: Generate C# code that opens an Excel file using Aspose.Cells, traverses each worksheet, sets the IsLocked property of every shape to true, and then exports the workbook to PDF. | Update an existing Aspose.Cells .NET routine to batch‑lock all drawing objects on each sheet prior to PDF conversion.
// Common Searches: asp.net lock all shapes in workbook using Aspose.Cells before PDF conversion | c# Aspose.Cells batch lock drawing objects on each worksheet | how to set shape.IsLocked for every shape in an Excel file with Aspose.Cells | export Excel workbook to PDF after locking shapes Aspose.Cells
// Tags: batch lock shapes Aspose.Cells | shape.IsLocked usage C# | traverse worksheets shape collection Aspose.Cells | save workbook as PDF Aspose.Cells | lock drawing objects Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through all worksheets, locks every shape by setting its IsLocked property to true, and then saves the workbook as a PDF, including basic error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Lock each shape
                foreach (Shape shape in shapes)
                {
                    shape.IsLocked = true; // Correct property to lock a shape
                }
            }

            // Export the workbook to PDF after all shapes have been locked
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
