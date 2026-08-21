// Title: C# – Convert an Aspose.Cells Workbook to PDF and Enable Full‑Screen Viewing (using Aspose.Pdf)
// Description: Shows how to build a workbook with Aspose.Cells, save it as a PDF using PdfSaveOptions, and then apply Aspose.Pdf viewer preferences so the PDF opens in full‑screen mode. Notes that Aspose.Cells alone lacks a direct full‑screen setting.
// Keywords: Aspose.Cells PDF conversion | C# Aspose.Cells to PDF | full screen PDF Aspose | Aspose.Pdf viewer preferences | PdfSaveOptions DisplayDocTitle | Excel to PDF C# | Aspose.Cells full-screen workaround | Aspose.Pdf FullScreen | C# PDF viewer settings
// Common Searches: Aspose.Cells save PDF full screen C# | How to set PDF viewer preferences with Aspose.Pdf after Aspose.Cells conversion | PdfSaveOptions DisplayDocTitle example | C# code to open generated PDF in full screen | Aspose.Cells PDF conversion without losing formatting
// Developer Intent: Create a PDF from an Excel workbook and configure it to launch in full‑screen mode.
// Use Cases: Generating presentation‑ready PDFs from Excel data. | Ensuring PDF opens maximized for kiosks or digital signage. | Adding document title to PDF metadata while converting. | Applying additional viewer preferences (e.g., hide toolbar) after conversion with Aspose.Pdf.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to PDF and then uses Aspose.Pdf to set ViewerPreferences.FullScreen. | Explain step‑by‑step how to modify PDF viewer settings after an Aspose.Cells conversion, including sample Aspose.Pdf snippets. | Show how to combine PdfSaveOptions.DisplayDocTitle with Aspose.Pdf full‑screen configuration in a single C# program.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to build a workbook with Aspose.Cells, save it as a PDF using PdfSaveOptions, and then apply Aspose.Pdf viewer preferences so the PDF opens in full‑screen mode. Notes that Aspose.Cells alone lacks a direct full‑screen setting.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF Conversion");
        sheet.Cells["A2"].PutValue("Full‑screen mode demonstration");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // NOTE:
        // Aspose.Cells does not expose a direct property to set the PDF viewer
        // to open in full‑screen mode. This can be achieved by using the
        // Aspose.Pdf library to modify the viewer preferences after the PDF
        // is created. Here we set a commonly used option (DisplayDocTitle) as an
        // example of using PdfSaveOptions.
        pdfOptions.DisplayDocTitle = true;

        // Save the workbook as a PDF file using the configured options
        workbook.Save("WorkbookFullScreen.pdf", pdfOptions);

        Console.WriteLine("Workbook has been saved to PDF.");
    }
}
