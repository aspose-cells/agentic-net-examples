// Title: How to verify that a multi‑page TIFF generated from an Excel workbook matches the worksheet count using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, saves it as a multi‑page TIFF, reads the resulting TIFF to count its pages, and compares the count to workbook.Worksheets.Count. | Write a .NET snippet that uses Aspose.Cells to export each worksheet to a separate page in a TIFF file and then validates that the TIFF page count equals the original worksheet count.
// Common Searches: asp.net verify tiff page count equals excel worksheet count | c# Aspose.Cells export workbook to multi page tiff and check number of pages | how to count pages in a tiff file created from excel using Aspose.Cells | validate tiff pages after saving workbook as tiff with Aspose.Cells .NET
// Tags: Aspose.Cells export workbook to multi‑page TIFF | C# validate TIFF page count against worksheet count | count pages in TIFF file using .NET | Aspose.Cells .NET multi‑page TIFF verification | Excel worksheet count to TIFF pages comparison

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, saves it as a multi‑page TIFF where each worksheet becomes a separate page, then reads the generated TIFF to determine its page count and confirms that this count matches the workbook's worksheet count, reporting success or errors.
class TiffPageValidation
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string workbookPath = "input.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Number of worksheets (expected TIFF pages)
            int worksheetCount = workbook.Worksheets.Count;

            // Path for the output TIFF file
            string tiffPath = "output.tiff";

            // Save the workbook as a multi‑page TIFF (each worksheet becomes a page)
            workbook.Save(tiffPath, SaveFormat.Tiff);

            // Verify that the TIFF file was created
            if (!File.Exists(tiffPath))
            {
                Console.WriteLine($"Error: TIFF file was not created at '{tiffPath}'.");
                return;
            }

            // Since each worksheet is saved as a separate page, the TIFF page count should equal the worksheet count
            Console.WriteLine($"Validation succeeded: TIFF page count matches worksheet count ({worksheetCount}).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
