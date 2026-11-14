// Title: Validate that all slicers are printable before saving an Excel workbook as PDF with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells check slicer IsPrintable property before PDF conversion | how to enforce printable slicers when exporting workbook to PDF using Aspose.Cells | prevent PDF export if any slicer is not set to printable in Aspose.Cells .NET | validate slicer printable flag in each worksheet prior to saving as PDF | Aspose.Cells .NET verify slicer printable status before workbook PDF save
// Tags: slicer printable flag validation Aspose.Cells | export workbook to PDF after slicer check | Aspose.Cells verify slicer IsPrintable C# | prevent PDF conversion non‑printable slicer | worksheet slicer printable property enforcement

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an Excel workbook, iterates through every worksheet's slicers to ensure the IsPrintable flag is true, throws an exception if a non‑printable slicer is found, and only then saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputFile);

            // Validate that every slicer in the workbook has its printable flag set to true
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Slicer slicer in sheet.Slicers)
                {
                    if (!slicer.IsPrintable)
                    {
                        throw new InvalidOperationException(
                            $"Slicer \"{slicer.Name}\" on sheet \"{sheet.Name}\" is not printable.");
                    }
                }
            }

            // Export the workbook to PDF after validation succeeds
            workbook.Save(outputFile, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputFile}\".");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"Validation error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
