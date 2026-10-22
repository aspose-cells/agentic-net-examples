// Title: Validate slicer printable flag before exporting Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over every worksheet in a Workbook, checks each slicer's IsPrintable property, and throws an InvalidOperationException if any slicer is not printable before calling Workbook.Save with SaveFormat.Pdf. | Create a helper method that uses dynamic typing to safely read the IsPrintable flag of slicer objects, logs the slicer name when the flag is false, and aborts the PDF conversion process. | Enhance an existing Aspose.Cells PDF export routine by adding a pre‑export validation step that ensures all slicers are marked printable and returns a detailed error message if validation fails.
// Common Searches: Aspose.Cells C# verify slicer IsPrintable property before PDF conversion | how to prevent PDF export when a slicer is not printable in Aspose.Cells | check all slicers printable flag in .NET workbook prior to saving as PDF | C# code sample for slicer printable validation using Aspose.Cells | exception handling for non‑printable slicers during Aspose.Cells PDF export
// Tags: slicer printable validation Aspose.Cells | export workbook to PDF with slicer checks | Aspose.Cells IsPrintable property C# | dynamic slicer property access .NET | prevent PDF generation on non‑printable slicer | exception handling for slicer printability

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Loads an Excel file, iterates through each worksheet's slicers using dynamic typing, validates that every slicer's IsPrintable property is true (throws an exception if not), and then saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Validate that every slicer's printable flag is true (using dynamic to avoid direct Slicer type reference)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // The Slicers collection may be empty; iterate if present
                foreach (var slicerObj in sheet.Slicers)
                {
                    try
                    {
                        dynamic slicer = slicerObj;
                        if (slicer.IsPrintable == false)
                        {
                            throw new InvalidOperationException(
                                $"Slicer '{slicer.Name}' on worksheet '{sheet.Name}' has IsPrintable set to false.");
                        }
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        // If the slicer object does not expose IsPrintable, skip validation
                    }
                }
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
