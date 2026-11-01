// Title: Export an Aspose.Cells workbook to PDF with all columns fitted on a single page per worksheet in C#
// Common Searches: Aspose.Cells C# export worksheet to PDF with all columns fitting on one page | Set AllColumnsInOnePagePerSheet to true in PdfSaveOptions for PDF conversion | C# code sample for fitting all columns on one PDF page per sheet using Aspose.Cells | How to compress worksheet columns to a single PDF page with Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet example for C# developers
// Tags: Aspose.Cells PdfSaveOptions column scaling | C# export workbook to PDF single-page columns | AllColumnsInOnePagePerSheet PDF conversion | worksheet column fit to single PDF page | Aspose.Cells PDF page layout settings

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a Workbook, fills it with data, configures PdfSaveOptions with AllColumnsInOnePagePerSheet set to true, and saves the workbook as a PDF so that each worksheet's columns are compressed to fit on one page.
class Program
{
    static void Main()
    {
        // Create a new workbook (replace with load logic if needed)
        Workbook workbook = new Workbook(); // create-workbook rule

        // OPTIONAL: populate some data for demonstration
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        for (int row = 0; row < 100; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Set PDF save options to fit all columns on one page per sheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions(); // create-pdfsaveoptions rule
        pdfOptions.AllColumnsInOnePagePerSheet = true;   // enable the required setting

        // Export the workbook to PDF using the configured options
        workbook.Save("AllColumnsOnePage.pdf", pdfOptions); // save-pdf rule
    }
}
