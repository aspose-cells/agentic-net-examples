// Title: C# – Center a header on every PDF page using Aspose.Cells
// Description: Shows how to add a centered header (Arial Bold 12 pt) to each page when exporting an Aspose.Cells workbook to PDF, with optional margin alignment via IsHFAlignMargins.
// Keywords: Aspose.Cells C# PDF header | SetHeader center Aspose.Cells | PdfSaveOptions header alignment | IsHFAlignMargins | center header each page | Excel to PDF header formatting | Aspose.Cells page setup header
// Common Searches: center header on each PDF page Aspose.Cells C# | Aspose.Cells SetHeader example | align header margins with page margins Aspose.Cells | export Excel to PDF with custom header Aspose | C# code to add header to PDF using Aspose.Cells
// Developer Intent: Add a centered header that appears at the top of every PDF page generated from an Aspose.Cells workbook.
// Use Cases: Place a company logo or name as a centered header on all pages of a financial report PDF. | Create a styled title header for multi‑sheet workbooks exported to PDF. | Maintain consistent header placement by aligning it with the page margins during PDF conversion.
// AI Prompts: Write C# code with Aspose.Cells to set a centered header (custom font) and export the workbook to PDF. | Explain how to switch the header section from left, center, or right using SetHeader in Aspose.Cells. | Show how IsHFAlignMargins affects header positioning when saving an Excel workbook as a PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to add a centered header (Arial Bold 12 pt) to each page when exporting an Aspose.Cells workbook to PDF, with optional margin alignment via IsHFAlignMargins.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data (optional, just to have content)
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["A2"].PutValue("More Data");

        // Set the header text in the center section (section index 1)
        // Use font commands to specify font name, style and size if needed
        // Example: Arial, Bold, size 12
        worksheet.PageSetup.SetHeader(1, "&\"Arial,Bold\"&12 My Header Text");

        // Optional: align header margins with page margins (default is true)
        worksheet.PageSetup.IsHFAlignMargins = true;

        // Save the workbook as PDF with default options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("HeaderCentered.pdf", pdfOptions);
    }
}
