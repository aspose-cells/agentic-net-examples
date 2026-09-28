// Title: Export selected worksheet indices to a single PDF using Aspose.Cells PdfSaveOptions.SheetSet in C#
// AI Prompts: Write C# code that loads an Excel file, creates a PdfSaveOptions object, assigns a SheetSet with an int[] of zero‑based worksheet indices, and saves the chosen sheets as one PDF with Aspose.Cells. | Show how to configure PdfSaveOptions.SheetSet to include only specific worksheets (e.g., 0 and 2) before calling Workbook.Save for PDF output. | Include a file‑existence check and exception handling while exporting the selected sheets to a PDF using Aspose.Cells in C#.
// Common Searches: Aspose.Cells how to export only certain worksheets to one PDF file in C# | PdfSaveOptions SheetSet set worksheet indices example | C# export Excel sheets 0 and 2 as a single PDF using Aspose.Cells | select multiple worksheets for PDF conversion with Aspose.Cells PdfSaveOptions
// Tags: Aspose.Cells PDF export selected worksheets | SheetSet int array configuration for PDF | C# zero‑based worksheet index selection | single PDF from multiple Excel sheets | file existence check Aspose.Cells export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads input.xlsx, creates PdfSaveOptions, sets SheetSet with worksheet indices 0 and 2, saves them as selected_sheets.pdf, and includes file‑existence validation and exception handling.
class ExportSelectedSheetsToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "selected_sheets.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define the set of worksheet indices to export (zero‑based)
            // Use the int[] constructor to avoid ambiguity with the string[] overload
            SheetSet sheetSet = new SheetSet(new int[] { 0, 2 });
            pdfOptions.SheetSet = sheetSet;

            // Export the selected sheets as a single PDF file
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Selected sheets have been exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
