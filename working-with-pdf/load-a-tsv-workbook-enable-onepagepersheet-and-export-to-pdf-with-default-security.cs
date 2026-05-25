using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the TSV workbook from file
        Workbook workbook = new Workbook("input.tsv");

        // Create PDF save options and enable OnePagePerSheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true; // Ensures each sheet fits on a single PDF page

        // Save the workbook as PDF using the options (default security – no passwords set)
        workbook.Save("output.pdf", pdfOptions);
    }
}