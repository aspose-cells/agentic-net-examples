// Title: Convert a TSV workbook to PDF with one page per sheet and no blank pages using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert TSV to PDF one page per worksheet | How to remove empty worksheets before saving PDF with Aspose.Cells | PdfSaveOptions OnePagePerSheet and AllColumnsInOnePagePerSheet example in C# | Generate PDF from TSV data without blank pages using Aspose.Cells | Load TSV file into Aspose.Cells workbook and export to PDF
// Tags: Aspose.Cells TSV workbook export | PdfSaveOptions single sheet per page | skip blank sheets before PDF conversion | fit all columns on one PDF page | avoid empty PDF pages

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads a TSV file into an Aspose.Cells Workbook, removes any worksheets that contain no data, configures PdfSaveOptions to place each sheet on a single PDF page with all columns fitting, and saves the result as a PDF without generating blank pages.
class TsvToPdfConverter
{
    static void Main()
    {
        // Paths for input TSV and output PDF
        string inputPath = "input.tsv";
        string outputPath = "output.pdf";

        // Load the TSV file into a workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Remove empty worksheets to prevent blank pages in the PDF
        for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
        {
            Worksheet sheet = workbook.Worksheets[i];
            // MaxDataRow/MaxDataColumn are -1 when the sheet has no data
            if (sheet.Cells.MaxDataRow < 0 || sheet.Cells.MaxDataColumn < 0)
            {
                workbook.Worksheets.RemoveAt(i);
            }
        }

        // Configure PDF save options: one page per sheet, all columns on that page
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true,
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as a PDF file
        workbook.Save(outputPath, pdfOptions);
    }
}
