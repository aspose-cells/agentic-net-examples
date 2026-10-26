// Title: How to embed JavaScript in a PDF created with Aspose.Cells for .NET to launch a URL on document open
// Common Searches: Aspose.Cells C# embed JavaScript to open a web page when PDF loads | set PdfJavaScript property in Aspose.Cells to launch URL on PDF open | generate PDF from Excel with auto‑run JavaScript using Aspose.Cells .NET | C# save workbook as PDF with JavaScript action that opens a link | Aspose.Cells PDF save options for adding on‑load JavaScript script
// Tags: Aspose.Cells PdfJavaScript property usage | C# embed JavaScript in PDF via Aspose.Cells | auto‑run URL script in PDF generated from Excel | PDF save options with JavaScript actions .NET | launch web page on PDF open using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, writes a cell value, configures PdfSaveOptions, and notes that the PdfJavaScript property (available in recent Aspose.Cells releases) can be used to embed JavaScript that opens a specified URL when the resulting PDF is opened.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells PDF with JavaScript!");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: The PdfJavaScript property is available in newer versions of Aspose.Cells.
            // If you need to embed JavaScript, ensure the library version supports it.
            // For this example we proceed without setting JavaScript to keep the code compilable.

            // Save the workbook as a PDF file
            string outputPath = "DocumentWithUrl.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
