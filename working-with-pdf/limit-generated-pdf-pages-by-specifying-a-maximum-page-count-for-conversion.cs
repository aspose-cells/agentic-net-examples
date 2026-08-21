// Title: C# – Cap PDF Pages When Converting Excel with Aspose.Cells
// Description: Shows how to load an .xlsx file, set Aspose.Cells PdfSaveOptions.PageCount (and optionally PageIndex) to restrict the PDF length, then save the workbook as a PDF in .NET.
// Keywords: Aspose.Cells | C# | PdfSaveOptions | PageCount | Excel to PDF | cap PDF pages | page range export | PDF conversion .NET | Workbook.Save PDF | Aspose.Cells PDF options
// Common Searches: Aspose.Cells cap PDF pages C# | PdfSaveOptions PageCount example | Export only first N pages from Excel to PDF | Set maximum PDF pages with Aspose.Cells | C# export Excel workbook to PDF with page range
// Developer Intent: Restrict the number of pages generated during Excel‑to‑PDF conversion using Aspose.Cells.
// Use Cases: Create a short preview PDF containing the initial pages of a large workbook | Produce a lightweight report for email by reducing PDF length | Enforce printing or storage limits by exporting only a set number of pages | Generate a specific page range for documentation or compliance purposes
// AI Prompts: Write a C# snippet that opens an .xlsx file, sets PdfSaveOptions.PageCount to a user‑defined value, and saves the workbook as a PDF with Aspose.Cells. | Explain how to combine PdfSaveOptions.PageIndex and PageCount to export a custom page range from an Excel workbook. | Show how to retrieve the total page count of a workbook, then limit the PDF output to that count or a lower number dynamically.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPageLimitDemo
{
    // Shows how to load an .xlsx file, set Aspose.Cells PdfSaveOptions.PageCount (and optionally PageIndex) to restrict the PDF length, then save the workbook as a PDF in .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Initialize PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the maximum number of pages to be saved.
            // For example, limit the output to the first 5 pages.
            pdfOptions.PageCount = 5;

            // Optionally, you can also set the starting page index (default is 0)
            // pdfOptions.PageIndex = 0;

            // Save the workbook to PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("PDF saved with a maximum of 5 pages.");
        }
    }
}
