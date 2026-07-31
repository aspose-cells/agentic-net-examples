// Title: Add Workbook Name as Center Header on Every PDF Page with Aspose.Cells for .NET (C#)
// Description: Shows how to place the workbook's file name—or a custom name—into a centered header on each page when exporting to PDF using Aspose.Cells PageSetup and PdfSaveOptions in C#.
// Keywords: Aspose.Cells PDF header | C# add workbook name to PDF | PdfSaveOptions header | center header &F placeholder | Aspose.Cells set header | display workbook title in PDF | Aspose.Cells .NET PDF export | PageSetup SetHeader example | workbook filename placeholder | PDF page header Aspose
// Common Searches: Aspose.Cells set PDF header to file name | C# add workbook name to each PDF page | How to use &F placeholder in Aspose.Cells | PdfSaveOptions display document title | Add centered header when exporting Excel to PDF | Show workbook title in PDF header Aspose
// Developer Intent: Insert the workbook’s name as a centered header on every page of the generated PDF.
// Use Cases: Creating multi‑page reports where each page displays the report title. | Exporting unsaved workbooks while preserving a meaningful name in the PDF header. | Automating branding of PDF documents generated from Excel files. | Ensuring consistent header information across all pages without post‑processing.
// AI Prompts: Generate C# code with Aspose.Cells that sets a custom workbook name and uses &F to place it in the center header of each PDF page. | Explain the role of PageSetup.SetHeader sections and the &F placeholder, and how PdfSaveOptions.DisplayDocTitle affects the PDF viewer. | Show how to assign Workbook.FileName for unsaved workbooks so the header displays the intended name. | Provide a step‑by‑step guide to export an Excel workbook to PDF with a centered header containing the workbook title.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to place the workbook's file name—or a custom name—into a centered header on each page when exporting to PDF using Aspose.Cells PageSetup and PdfSaveOptions in C#.
class AddWorkbookNameHeaderToPdf
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a page header that displays the file name (without path)
        // Section 1 = center section of the header
        // &F inserts the file name; if the workbook is not saved yet,
        // Aspose.Cells uses the workbook name set in the FileName property.
        sheet.PageSetup.SetHeader(1, "&F");

        // Optionally set the workbook name (used when the workbook is not saved to disk)
        // This ensures the header shows a meaningful name.
        workbook.FileName = "MyWorkbook";

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Display the document title in the PDF viewer's title bar (optional)
            DisplayDocTitle = true
        };

        // Save the workbook as PDF; the header will appear on every page
        string outputPath = "WorkbookWithHeader.pdf";
        workbook.Save(outputPath, pdfOptions);

        Console.WriteLine($"PDF saved to {Path.GetFullPath(outputPath)}");
    }
}
