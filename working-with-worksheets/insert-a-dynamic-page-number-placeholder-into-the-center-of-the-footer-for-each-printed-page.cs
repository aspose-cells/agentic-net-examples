// Title: Add "Page X of Y" to the center footer of an Excel sheet with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, optionally adds sample data, and uses Worksheet.PageSetup.SetFooter to insert the "Page &P of &N" placeholder in the center section of the printed footer, then saves the file as FooterWithPageNumber.xlsx.
// Keywords: Aspose.Cells C# set footer | Excel page number placeholder | center footer page count | Worksheet PageSetup SetFooter | dynamic page numbering Aspose | print footer Aspose.Cells
// Common Searches: Aspose.Cells set center footer page number | C# add page X of Y to Excel footer | How to use &P &N in Aspose.Cells footer | Print Excel with page numbers using Aspose.Cells .NET | Set footer for all worksheets Aspose.Cells
// Developer Intent: Place a dynamic "Page X of Y" placeholder in the center of the worksheet footer.
// Use Cases: Generate multi‑page printable reports where each page shows "Page X of Y" in the footer. | Create invoices or statements that automatically include page numbering for proper pagination. | Produce catalogs, manuals, or documentation exported from Excel with consistent footer pagination. | Automate batch export of workbooks to PDF/print with uniform page numbering across sheets.
// AI Prompts: Generate C# code to set left and right footer text while keeping the center page number placeholder using Aspose.Cells. | Show how to apply the same footer settings to every worksheet in a workbook with Aspose.Cells. | Explain how to format the page number placeholder in the footer (e.g., bold, font size) using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFooterExample
{
    // Creates a workbook, optionally adds sample data, and uses Worksheet.PageSetup.SetFooter to insert the "Page &P of &N" placeholder in the center section of the printed footer, then saves the file as FooterWithPageNumber.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data (optional, just to have printable content)
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["A2"].PutValue("More Data");

            // Set the center section of the footer to display the current page number and total pages
            // &P = current page number, &N = total page count
            worksheet.PageSetup.SetFooter(1, "Page &P of &N");

            // Save the workbook to a file
            workbook.Save("FooterWithPageNumber.xlsx");
        }
    }
}
