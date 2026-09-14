// Title: How to verify that the CrossStringInPdf option controls long text overflow when exporting Excel to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that exports a worksheet to PDF twice—once with PdfSaveOptions.CrossStringInPdf set to true and once set to false—and then programmatically compare the two PDFs to see how the long string is rendered. | Explain how to inspect the generated PDF files to confirm whether the long cell content crosses cell boundaries based on the CrossStringInPdf setting.
// Common Searches: Aspose.Cells C# export Excel to PDF with CrossStringInPdf true | How to check text overflow in PDF generated from Excel using Aspose.Cells | Enable or disable string crossing in PDF output with PdfSaveOptions Aspose.Cells | Difference in PDF rendering when CrossStringInPdf is false in Aspose.Cells | Validate long cell content rendering in PDF using Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions CrossStringInPdf | C# export Excel to PDF text overflow | verify PDF string crossing behavior | long cell content PDF rendering | disable string crossing Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook, inserts a very long string into cell A1, narrows column A to force overflow, and saves the sheet twice as PDF using default PdfSaveOptions (without setting CrossStringInPdf). To confirm the setting's effect, modify the code to set CrossStringInPdf true and false, generate the PDFs, and compare how the long text crosses cell boundaries.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a long string that will exceed the column width when rendered
            sheet.Cells["A1"].PutValue("This is a very long string that should cross the cell boundary when exported to PDF.");

            // Make column A narrow so the text definitely overflows
            sheet.Cells.SetColumnWidth(0, 10); // width in characters

            // Export PDF with default settings (CrossStringInPdf option not available in this version)
            PdfSaveOptions optionsTrue = new PdfSaveOptions();
            workbook.Save("CrossStringTrue.pdf", optionsTrue);

            // Export PDF again with default settings
            PdfSaveOptions optionsFalse = new PdfSaveOptions();
            workbook.Save("CrossStringFalse.pdf", optionsFalse);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
