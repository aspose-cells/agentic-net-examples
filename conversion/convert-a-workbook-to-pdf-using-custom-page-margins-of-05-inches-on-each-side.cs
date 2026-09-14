// Title: Convert an Excel workbook to PDF with 0.5‑inch page margins using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, sets all page margins to 0.5 inches via PageSetup, and saves it as a PDF with Aspose.Cells. | Show how to configure left, right, top, and bottom margins in inches before exporting an Excel sheet to PDF in a .NET application. | Provide a minimal Aspose.Cells example that applies half‑inch margins to a worksheet and generates a PDF file. | Demonstrate setting PageSetup.MarginInch properties and invoking Workbook.Save with SaveFormat.Pdf.
// Common Searches: Aspose.Cells how to set 0.5 inch margins before PDF export in C# | C# export Excel to PDF with custom page margins using Aspose.Cells | set worksheet page margins inches Aspose.Cells .NET example | convert workbook to PDF with specific margins Aspose.Cells | adjust PDF output margins in Aspose.Cells PageSetup
// Tags: Aspose.Cells PDF export custom margins | C# set worksheet page margins inches | Aspose.Cells PageSetup margin configuration | convert Excel to PDF half‑inch margins | Aspose.Cells SaveFormat.Pdf with margins

using System;
using Aspose.Cells;

// The program creates a new workbook, configures the left, right, top, and bottom margins to 0.5 inches via the PageSetup properties, adds optional sample data, and saves the workbook as a PDF file named WorkbookWithMargins.pdf using Aspose.Cells.
class ConvertWorkbookToPdfWithMargins
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set custom page margins of 0.5 inches on each side
        worksheet.PageSetup.LeftMarginInch = 0.5;
        worksheet.PageSetup.RightMarginInch = 0.5;
        worksheet.PageSetup.TopMarginInch = 0.5;
        worksheet.PageSetup.BottomMarginInch = 0.5;

        // Add sample data (optional, just to have visible content)
        worksheet.Cells["A1"].PutValue("Workbook converted to PDF with 0.5 inch margins.");

        // Save the workbook as a PDF file
        workbook.Save("WorkbookWithMargins.pdf", SaveFormat.Pdf);
    }
}
