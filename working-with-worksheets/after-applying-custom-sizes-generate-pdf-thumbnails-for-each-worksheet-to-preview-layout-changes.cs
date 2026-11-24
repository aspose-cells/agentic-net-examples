// Title: Generate PDF thumbnail previews for each worksheet after applying custom column widths and row heights using Aspose.Cells for .NET
// Common Searches: how to create a PDF preview of each worksheet after setting column width with Aspose.Cells C# | save each Excel sheet as a single-page PDF thumbnail using Aspose.Cells .NET | apply custom row height and column width then export worksheet to PDF preview in C# | Aspose.Cells generate per-sheet PDF thumbnail with zoom 50 percent
// Tags: custom column width Aspose.Cells C# | custom row height Aspose.Cells C# | single-page PDF export Aspose.Cells | worksheet PDF thumbnail generation Aspose.Cells | PdfSaveOptions OnePagePerSheet Aspose.Cells | temporary workbook copy worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, sets column A to 20 characters and row 1 to 30 points on every worksheet, then iterates through each sheet, reduces its page zoom to 50 % to act as a thumbnail, copies the sheet into a temporary workbook, configures PdfSaveOptions with OnePagePerSheet, and saves each sheet as a separate "{SheetName}_preview.pdf" file while handling sheet‑specific and global exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // Apply custom column widths and row heights
            // -------------------------------------------------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Example: set column A (index 0) width to 20 characters
                sheet.Cells.SetColumnWidth(0, 20);

                // Example: set row 1 (index 0) height to 30 points
                sheet.Cells.SetRowHeight(0, 30);
            }

            // -------------------------------------------------
            // Generate a PDF preview (thumbnail) for each worksheet
            // -------------------------------------------------
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                try
                {
                    Worksheet sourceSheet = workbook.Worksheets[i];

                    // Reduce the zoom to make the PDF act as a thumbnail preview
                    sourceSheet.PageSetup.Zoom = 50; // 50% of original size

                    // Create a temporary workbook containing only the current sheet
                    Workbook tempWorkbook = new Workbook();
                    tempWorkbook.Worksheets.Clear();

                    // Add a copy of the source sheet by name
                    tempWorkbook.Worksheets.AddCopy(sourceSheet.Name);

                    // Configure PDF save options
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        OnePagePerSheet = true // keep the sheet on a single page
                    };

                    // Build a file name for the thumbnail PDF
                    string pdfFileName = $"{sourceSheet.Name}_preview.pdf";

                    // Save the single‑sheet PDF using the options overload
                    tempWorkbook.Save(pdfFileName, pdfOptions);
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Failed to generate PDF for sheet index {i}: {exSheet.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
