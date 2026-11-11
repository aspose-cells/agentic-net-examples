// Title: Export only the first ten pages of an Excel workbook to PDF using Aspose.Cells PdfSaveOptions in C#
// Common Searches: Aspose.Cells C# export Excel to PDF limited to first 10 pages | Set PdfSaveOptions.PageCount to restrict PDF output in Aspose.Cells | How to save a workbook as PDF with a maximum of ten pages using Aspose.Cells | C# code to limit PDF conversion to ten pages with Aspose.Cells | Aspose.Cells PDF conversion page count example in .NET
// Tags: Aspose.Cells PdfSaveOptions page count C# | limit PDF pages during Excel to PDF conversion | export Excel workbook to PDF first ten pages | Workbook.Save with PdfSaveOptions page limit | C# restrict PDF output pages Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures PdfSaveOptions by setting PageCount to 10 to limit the PDF output to the first ten pages, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to limit output to the first ten pages
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.PageCount = 10; // limit to first ten pages

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
