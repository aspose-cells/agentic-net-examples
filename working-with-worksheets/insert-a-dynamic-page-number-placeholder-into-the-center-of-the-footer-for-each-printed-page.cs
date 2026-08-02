// Title: Add Centered "Page X of Y" Footer in Excel using Aspose.Cells for C#
// Description: Demonstrates how to set the center section of an Excel worksheet footer to show the current page number and total pages ("Page &P of &N") with Aspose.Cells, then save the workbook.
// Keywords: Aspose.Cells C# footer page number | Excel center footer page count | SetFooter Aspose.Cells example | dynamic page number placeholder &P &N | C# add page numbers to printed Excel
// Common Searches: Aspose.Cells set center footer page number | C# add "Page X of Y" to Excel footer | How to use SetFooter for page numbering in Aspose.Cells | Print Excel with page numbers using Aspose.Cells
// Developer Intent: Insert a placeholder that automatically displays the current page and total page count in the center footer of every printed worksheet.
// Use Cases: Printing multi‑page reports that need a consistent "Page X of Y" label. | Generating invoices or statements where each sheet shows its position in the document. | Automating documentation exports that require centered page numbering for compliance.
// AI Prompts: Show C# code to set left, center, and right footer sections with different Aspose.Cells placeholders. | Explain how to change the font style and size of footer text in Aspose.Cells. | Provide a method to hide the footer on the first printed page while keeping page numbers on subsequent pages.

using System;
using Aspose.Cells;

namespace AsposeCellsFooterExample
{
    // Demonstrates how to set the center section of an Excel worksheet footer to show the current page number and total pages ("Page &P of &N") with Aspose.Cells, then save the workbook.
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
