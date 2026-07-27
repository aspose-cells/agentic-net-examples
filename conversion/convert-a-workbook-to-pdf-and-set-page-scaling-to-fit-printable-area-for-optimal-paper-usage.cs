// Title: C# – Convert Excel Workbook to PDF with Fit‑to‑Page Width Scaling using Aspose.Cells
// Description: Creates a workbook, populates the first worksheet with sample data, applies PageSetup.SetFitToPages(1, 0) to fit the printable area to one page wide (height auto‑adjusts), and saves the file as a PDF. Demonstrates optimal paper usage without manual scaling.
// Keywords: Aspose.Cells | C# PDF conversion | SetFitToPages | fit to page width | Excel to PDF | page scaling | printable area | SaveFormat.Pdf
// Common Searches: Aspose.Cells set fit to page when saving PDF | C# convert Excel to PDF fit width | How to use SetFitToPages in Aspose.Cells | Fit printable area PDF Aspose.Cells .NET | Scale Excel sheet to one page width PDF
// Developer Intent: Export an Excel workbook to PDF while automatically scaling the sheet to fit the printable width (one page) for efficient paper usage.
// Use Cases: Generate invoices that print on a single page width without manual adjustments. | Produce multi‑sheet financial reports where each sheet is uniformly scaled to one page wide. | Batch‑process a folder of workbooks into PDFs with consistent fit‑to‑page scaling for archiving.
// AI Prompts: Write C# code that converts an Aspose.Cells workbook to PDF using fit‑to‑page width scaling and custom margins. | Explain the parameters of SetFitToPages in Aspose.Cells and how they affect PDF output. | Provide a script to convert multiple Excel files to PDF, applying the same one‑page‑wide scaling to each workbook.

using System;
using Aspose.Cells;

// Creates a workbook, populates the first worksheet with sample data, applies PageSetup.SetFitToPages(1, 0) to fit the printable area to one page wide (height auto‑adjusts), and saves the file as a PDF. Demonstrates optimal paper usage without manual scaling.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        for (int row = 0; row < 30; row++)
        {
            sheet.Cells[row, 0].PutValue($"Item {row + 1}");
            sheet.Cells[row, 1].PutValue((row + 1) * 5);
        }

        // Set page scaling to fit the printable area:
        // Fit to 1 page wide and let the height adjust automatically (0 means auto)
        sheet.PageSetup.SetFitToPages(1, 0);

        // Convert the workbook to PDF using the default PDF save options
        workbook.Save("ConvertedWorkbook.pdf", SaveFormat.Pdf);

        Console.WriteLine("Workbook has been successfully converted to PDF with fit-to-page scaling.");
    }
}
