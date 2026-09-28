// Title: Generate a multi‑page TIFF from the first three worksheets of an Excel file using Aspose.Cells for .NET
// AI Prompts: Create a C# program that loads an Excel workbook, copies the first three worksheets into a new workbook, and saves the result as a multi‑page TIFF with Aspose.Cells. | Write code to export up to three selected sheets from a source workbook to a single TIFF image, each sheet becoming a separate page, using Aspose.Cells SaveFormat.Tiff.
// Common Searches: Aspose.Cells export first three Excel sheets to multi page TIFF in C# | How to save selected worksheets as a multi‑page TIFF image using .NET | C# code to convert specific Excel worksheets to a TIFF file with Aspose.Cells | Multi‑page TIFF generation from multiple worksheets using Aspose.Cells SaveFormat.Tiff
// Tags: export selected worksheets to multi-page TIFF Aspose.Cells | copy first three sheets to new workbook C# | Aspose.Cells SaveFormat.Tiff for multiple sheets | C# generate multi-page TIFF from Excel | limit worksheet copy to three sheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input.xlsx file, loads it into a Workbook, creates a new Workbook, removes its default sheet, copies up to the first three worksheets from the source workbook, and saves the new workbook as a multi‑page TIFF file named output.tiff using Aspose.Cells.
class MultiPageTiffGenerator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Create a new workbook to hold only the first three worksheets
            Workbook tiffWorkbook = new Workbook();

            // Remove the default empty worksheet that a new workbook contains
            if (tiffWorkbook.Worksheets.Count > 0)
                tiffWorkbook.Worksheets.RemoveAt(0);

            // Copy the first three worksheets (or fewer if the source has less)
            int sheetsToCopy = Math.Min(3, sourceWorkbook.Worksheets.Count);
            for (int i = 0; i < sheetsToCopy; i++)
            {
                // Add a copy of the worksheet to the new workbook using its name
                tiffWorkbook.Worksheets.AddCopy(sourceWorkbook.Worksheets[i].Name);
            }

            // Save the new workbook as a multi‑page TIFF.
            // Each worksheet will become a separate page in the TIFF file.
            tiffWorkbook.Save(outputPath, SaveFormat.Tiff);
            Console.WriteLine($"Multi‑page TIFF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
