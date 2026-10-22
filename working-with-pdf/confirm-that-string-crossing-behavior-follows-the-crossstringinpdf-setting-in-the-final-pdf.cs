// Title: Verify Aspose.Cells PDF export string crossing behavior using the CrossStringInPdf setting in C#
// AI Prompts: Write C# code that sets PdfSaveOptions.CrossStringInPdf = false, exports a worksheet containing a long string, and describes how to visually confirm that the text stays within its original cell in the resulting PDF. | Create a C# example that toggles PdfSaveOptions.CrossStringInPdf between true and false, generates two PDFs, and programmatically compares the rendered text positions to demonstrate the effect of the setting. | Provide step‑by‑step instructions for testing the CrossStringInPdf property by exporting a workbook to PDF, opening the file, and verifying whether the long text spans multiple cells.
// Common Searches: Aspose.Cells how to prevent text from crossing cells when exporting to PDF in .NET | C# Aspose.Cells PDF export string crossing setting example | Check string overflow in PDF generated from Excel using Aspose.Cells | Validate cell text layout after PDF conversion with Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions CrossStringInPdf | C# export Excel to PDF without text overflow | validate PDF text layout Aspose.Cells | control string crossing in PDF export | column width effect on PDF rendering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

// The sample creates a workbook, places a long string in cell A1, adjusts column widths, and saves the file as a PDF using PdfSaveOptions. By changing the CrossStringInPdf property, developers can observe whether the long text stays within its cell or crosses into adjacent cells, enabling verification of the setting's impact on PDF rendering.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a long string that will normally cross cell boundaries
            // This will help us observe the string crossing behavior in the exported PDF
            sheet.Cells["A1"].PutValue("This is a very long string that should cross multiple cells when rendered.");

            // Optionally, set column widths to make the crossing more evident
            sheet.Cells.SetColumnWidth(0, 20); // Column A
            sheet.Cells.SetColumnWidth(1, 20); // Column B
            sheet.Cells.SetColumnWidth(2, 20); // Column C

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            string pdfPath = "CrossStringDemo.pdf";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(pdfPath, pdfOptions);

            // Output confirmation
            Console.WriteLine($"Workbook exported to PDF at '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
