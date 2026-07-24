// Title: C# – Convert Excel to PDF without Cell Comments using AspNet Aspose.Cells
// Description: Loads an .xlsx workbook, clears every worksheet's comments, optionally removes personal metadata, applies default PdfSaveOptions, and saves a clean PDF file.
// Keywords: Aspose.Cells | C# | Excel to PDF conversion | remove cell comments | clear worksheet comments | PdfSaveOptions | remove personal information | export clean PDF | Aspose.Cells PDF export
// Common Searches: Aspose.Cells export Excel to PDF without comments | C# remove cell comments before PDF conversion | How to clear worksheet comments in Aspose.Cells | Generate PDF from Excel without personal data using Aspose | Save workbook as PDF excluding comments Aspose.Cells
// Developer Intent: Create a PDF from an Excel file while stripping all cell comments and optional personal metadata.
// Use Cases: Client‑facing reports that must hide internal reviewer notes. | Compliance documents where comments are prohibited by policy. | Printable manuals derived from spreadsheets without author remarks. | Automated batch conversion of workbooks to clean PDFs for archiving.
// AI Prompts: Provide C# code using Aspose.Cells that removes all comments from each worksheet and saves the workbook as a PDF. | Explain the steps to clear cell comments and delete personal information before exporting an Excel file to PDF with Aspose.Cells. | Show how to verify that no comments remain in a workbook prior to PDF generation using Aspose.Cells methods.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // Loads an .xlsx workbook, clears every worksheet's comments, optionally removes personal metadata, applies default PdfSaveOptions, and saves a clean PDF file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and clear their comments
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.ClearComments(); // Removes all cell comments from the sheet
            }

            // Optionally remove any remaining personal information (author names, etc.)
            workbook.RemovePersonalInformation();

            // Configure PDF save options (default options are sufficient for this task)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the cleaned workbook as a PDF
            string outputPath = "output_clean.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook converted to PDF without comments: {outputPath}");
        }
    }
}
