// Title: Lock all shapes on every worksheet before exporting the workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through each worksheet in a Workbook, sets Shape.IsLocked = true for every shape, and then saves the workbook as a PDF using Aspose.Cells. | Create a reusable C# method that receives an Excel file path, locks all shapes across all sheets, and outputs a PDF where the shapes are protected from editing.
// Common Searches: Aspose.Cells how to lock every shape in an Excel file before PDF conversion | C# batch lock shapes on all worksheets using Aspose.Cells | prevent editing of shapes when saving Excel as PDF with Aspose.Cells | iterate through worksheets and set Shape.IsLocked property in Aspose.Cells | export workbook to PDF after locking shapes Aspose.Cells .NET
// Tags: lock shapes Aspose.Cells | batch shape locking C# | export workbook to PDF Aspose.Cells | Shape.IsLocked property usage | iterate worksheets lock shapes

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// // Loads an Excel workbook, iterates each worksheet to set Shape.IsLocked = true for all shapes, then saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Lock all shapes on every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    shape.IsLocked = true; // Prevent editing of the shape
                }
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
