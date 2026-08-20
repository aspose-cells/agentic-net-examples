// Title: C# – Turn Off Row/Column Headings When Exporting an Aspose.Cells Workbook to PDF
// Description: The sample builds a new workbook, inserts a few product rows, sets PageSetup.PrintHeadings to false to keep headers out of the printed output, optionally hides the UI headers with IsRowColumnHeadersVisible, and saves the result as a PDF suitable for a clean marketing brochure.
// Keywords: Aspose.Cells | C# | PrintHeadings | HideHeaders | PDF export | worksheet headings | disable printing headings | Excel to PDF | marketing brochure | PageSetup
// Common Searches: Aspose.Cells hide row and column headings in PDF | C# disable printing of worksheet headers with Aspose.Cells | PageSetup.PrintHeadings false example | Export Excel to PDF without headers using Aspose | How to remove Excel grid headings in a brochure PDF
// Developer Intent: Exclude row and column headings from the printed/PDF version of a worksheet and optionally conceal them in the on‑screen view before saving.
// Use Cases: Create a product catalog PDF that shows only data rows, no Excel headers. | Generate a clean financial summary report for distribution as a brochure. | Produce a marketing flyer from Excel data where UI headers are hidden for a professional look.
// AI Prompts: Write C# code with Aspose.Cells to suppress row/column headings during PDF export. | Explain the difference between PageSetup.PrintHeadings and IsRowColumnHeadersVisible in Aspose.Cells. | Show how to hide worksheet headers both on screen and in the printed PDF for a brochure layout.

using System;
using Aspose.Cells;

namespace MarketingBrochure
{
    // The sample builds a new workbook, inserts a few product rows, sets PageSetup.PrintHeadings to false to keep headers out of the printed output, optionally hides the UI headers with IsRowColumnHeadersVisible, and saves the result as a PDF suitable for a clean marketing brochure.
    class DisableRowHeadingsPrint
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (optional, for demonstration)
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(2.5);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(1.8);

            // Disable printing of row and column headings
            worksheet.PageSetup.PrintHeadings = false;

            // (Optional) Hide row/column headers in the UI as well
            worksheet.IsRowColumnHeadersVisible = false;

            // Save the workbook (e.g., as PDF for brochure distribution)
            workbook.Save("Brochure.pdf", SaveFormat.Pdf);

            Console.WriteLine("Workbook saved with row headings disabled for printing.");
        }
    }
}
