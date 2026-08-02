// Title: C# – Set 0.5‑inch margins and fit a worksheet to a single page with Aspose.Cells
// Description: Shows how to create a workbook, assign half‑inch top, bottom, left and right margins via PageSetup, apply SetFitToPages(1,1) to scale the sheet to one printed page, and save the result as Output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | PageSetup margins | half‑inch margin | fit worksheet to one page | SetFitToPages | Excel printing | scale to single page
// Common Searches: Aspose.Cells set page margins in inches | C# fit Excel sheet to one page programmatically | How to use SetFitToPages with custom margins | Print Excel worksheet on one page using Aspose.Cells | Adjust margins before scaling worksheet Aspose.Cells
// Developer Intent: Configure half‑inch margins and scale a worksheet so it prints on a single page.
// Use Cases: Printing standardized reports with half‑inch margins on one sheet | Generating invoices that must fit on a single page while preserving layout | Exporting data to PDF with exact margin and page‑fit settings
// AI Prompts: Provide C# code to set margins in centimeters with Aspose.Cells. | Show how to combine custom margins with landscape orientation and A4 paper size. | Explain how to preserve aspect ratio when using SetFitToPages with different margin values.

using Aspose.Cells;

// Shows how to create a workbook, assign half‑inch top, bottom, left and right margins via PageSetup, apply SetFitToPages(1,1) to scale the sheet to one printed page, and save the result as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Apply 0.5 inch margins on all sides
        sheet.PageSetup.BottomMarginInch = 0.5;
        sheet.PageSetup.TopMarginInch = 0.5;
        sheet.PageSetup.LeftMarginInch = 0.5;
        sheet.PageSetup.RightMarginInch = 0.5;

        // Scale the worksheet to fit on a single page
        sheet.PageSetup.SetFitToPages(1, 1);

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
