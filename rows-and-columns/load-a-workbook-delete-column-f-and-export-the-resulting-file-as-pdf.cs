// Title: Delete column F from an Excel worksheet and export the workbook as a PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, removes column F from the first sheet, and saves the workbook directly to a PDF with one page per sheet using Aspose.Cells. | Generate a snippet that deletes a specific column in a worksheet, updates any dependent formulas, and applies PdfSaveOptions before exporting to PDF with Aspose.Cells. | Show how to combine Cells.DeleteColumn and PdfSaveOptions to convert an edited Excel file to a PDF in C#.
// Common Searches: asp.net delete column F from Excel and convert to PDF with Aspose.Cells | c# remove specific column before exporting workbook to PDF using Aspose.Cells | how to keep formula references after deleting a column in Aspose.Cells | Aspose.Cells PdfSaveOptions one page per sheet example | delete column in first worksheet and save as PDF in C#
// Tags: delete column Aspose.Cells C# | excel to pdf conversion Aspose.Cells | PdfSaveOptions one page per sheet | update formulas after column deletion Aspose.Cells | remove specific worksheet column before pdf export

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions if needed

// The example loads an existing Excel workbook, deletes column F (zero‑based index 5) from the first worksheet while preserving formula references, configures PdfSaveOptions to place each sheet on a single PDF page, and saves the modified workbook as a PDF file.
class DeleteColumnAndExportPdf
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete column F (zero‑based index 5) and update references in other worksheets
        sheet.Cells.DeleteColumn(5, true);

        // Create PDF save options (optional – you can customize as needed)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true // example option
        };

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
