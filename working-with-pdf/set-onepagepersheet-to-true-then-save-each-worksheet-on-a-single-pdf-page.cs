// Title: C# – Export Each Excel Worksheet to a Single‑Page PDF with Aspose.Cells
// Description: Shows how to load an Excel workbook, set PdfSaveOptions.OnePagePerSheet to true, use SheetSet to target each worksheet, and save every sheet as a separate one‑page PDF file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | PdfSaveOptions | OnePagePerSheet | SheetSet | Excel to PDF conversion | single page PDF | separate PDF per worksheet | batch PDF generation | export worksheet as PDF
// Common Searches: Aspose.Cells export each sheet to single page PDF | C# PdfSaveOptions OnePagePerSheet example | How to save Excel worksheets as individual PDFs with Aspose.Cells | SheetSet usage in Aspose.Cells PDF conversion | Generate one‑page PDF per worksheet C#
// Developer Intent: Create individual one‑page PDF files for all worksheets in an Excel workbook.
// Use Cases: Produce printable reports where each Excel sheet appears on its own PDF page. | Generate separate invoice PDFs from individual worksheets. | Automate batch conversion of a multi‑sheet workbook into distinct single‑page PDFs for distribution.
// AI Prompts: Write C# code that uses Aspose.Cells to convert every worksheet in an Excel file into a separate PDF with one page per sheet. | Explain how PdfSaveOptions.OnePagePerSheet and SheetSet work together to export specific worksheets as single‑page PDFs. | Suggest how to modify the sample to save PDFs into a custom directory and name them using the worksheet titles.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsOnePagePerSheetDemo
{
    // Shows how to load an Excel workbook, set PdfSaveOptions.OnePagePerSheet to true, use SheetSet to target each worksheet, and save every sheet as a separate one‑page PDF file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create PDF save options and enable OnePagePerSheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.OnePagePerSheet = true; // Ensure each sheet fits on a single PDF page

            // Iterate through all worksheets and save each one as a separate PDF file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Set the SheetSet to the current worksheet index
                pdfOptions.SheetSet = new SheetSet(new int[] { i });

                // Define output file name (e.g., Sheet1.pdf, Sheet2.pdf, ...)
                string outputFile = $"Sheet{i + 1}.pdf";

                // Save the workbook (only the selected sheet) to PDF
                workbook.Save(outputFile, pdfOptions);
            }

            Console.WriteLine("All worksheets have been saved as single-page PDFs.");
        }
    }
}
