// Title: Create a PDF from an Aspose.Cells workbook where each worksheet appears on its own page using PdfSaveOptions.OnePagePerSheet in C#
// AI Prompts: Write C# code that saves an Excel workbook to PDF with one PDF page per worksheet using Aspose.Cells. | Demonstrate how to enable the OnePagePerSheet flag in PdfSaveOptions before exporting a workbook to PDF in .NET. | Adapt the sample to load an existing .xlsx file and generate a PDF where each sheet is rendered on a separate page.
// Common Searches: Aspose.Cells C# PdfSaveOptions.OnePagePerSheet example | export each Excel worksheet to a separate PDF page using Aspose.Cells | how to force one PDF page per sheet when saving workbook to PDF in .NET | C# Aspose.Cells PDF pagination per worksheet
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# export Excel to PDF per worksheet | set OnePagePerSheet true Aspose.Cells | PDF pagination per worksheet Aspose.Cells | save workbook as PDF with sheet pagination

using Aspose.Cells;

// The code creates or loads a Workbook, sets PdfSaveOptions.OnePagePerSheet to true, and saves the workbook as a PDF so that each worksheet is rendered on a separate PDF page.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example: add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data");

        // Configure PDF save options to place each worksheet on a separate page
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true; // Force one page per sheet

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
