// Title: Convert an Excel workbook to a PDF and strip all form control shapes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove all form controls from Excel before PDF export | how to delete shapes in a worksheet with Aspose.Cells to create a non‑editable PDF | export Excel to PDF without editable fields using Aspose.Cells .NET | C# code to clear worksheet ShapeCollection then save as PDF with Aspose.Cells | remove form field objects from .xlsx and generate locked PDF via Aspose.Cells
// Tags: remove form control shapes Aspose.Cells | Excel to PDF conversion non‑editable Aspose.Cells | clear worksheet ShapeCollection C# | Aspose.Cells PDF export without form fields | C# strip Excel shapes before PDF generation

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an .xlsx file with Aspose.Cells, iterates through each worksheet to delete every shape in the ShapeCollection (removing form controls), and then saves the workbook as a PDF, producing a final document that cannot be edited via form fields.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and remove form control shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Remove shapes in reverse order to avoid index shifting
                for (int i = shapes.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        // Delete the shape (Aspose.Cells does not provide a Delete method;
                        // removal is done via the ShapeCollection)
                        shapes.RemoveAt(i);
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Failed to delete shape at index {i} on sheet '{sheet.Name}': {exShape.Message}");
                    }
                }
            }

            // Save the modified workbook as a PDF (non‑editable final version)
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
