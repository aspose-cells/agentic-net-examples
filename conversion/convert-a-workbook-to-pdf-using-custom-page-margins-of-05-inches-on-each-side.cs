// Title: C# – Convert Aspose.Cells Workbook to PDF with 0.5‑inch Page Margins
// Description: Shows how to set left, right, top, and bottom margins to 0.5 inches via Aspose.Cells PageSetup and then save the workbook as a PDF using SaveFormat.Pdf.
// Keywords: Aspose.Cells PDF margins | C# set page margins inches | Workbook.Save PDF Aspose.Cells | PageSetup margin Aspose.Cells | custom PDF margins C# | export Excel to PDF with margins
// Common Searches: Aspose.Cells set PDF margins C# | C# export Excel to PDF with 0.5 inch margins | how to change page margins before PDF conversion Aspose.Cells | PageSetup margin inches example | save workbook as PDF with custom margins
// Developer Intent: Generate a PDF from an Excel workbook while applying a uniform 0.5‑inch margin on every side.
// Use Cases: Produce printable reports that require half‑inch margins for a clean layout. | Export invoices or statements to PDF with consistent margin settings for branding. | Create handouts or forms where precise margin control is needed for binding or filing.
// AI Prompts: Provide C# code that sets 0.5‑inch margins and converts an Aspose.Cells workbook to PDF. | Explain how to adjust page margins in inches before saving a workbook as PDF with Aspose.Cells. | Show an example of error‑handled PDF export with custom margins using Aspose.Cells in C#.

using System;
using Aspose.Cells;

// Shows how to set left, right, top, and bottom margins to 0.5 inches via Aspose.Cells PageSetup and then save the workbook as a PDF using SaveFormat.Pdf.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set custom page margins of 0.5 inches on each side
        sheet.PageSetup.LeftMarginInch = 0.5;
        sheet.PageSetup.RightMarginInch = 0.5;
        sheet.PageSetup.TopMarginInch = 0.5;
        sheet.PageSetup.BottomMarginInch = 0.5;

        // Add sample data (optional, just to have content in the PDF)
        sheet.Cells["A1"].PutValue("Workbook converted to PDF with 0.5 inch margins.");

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
