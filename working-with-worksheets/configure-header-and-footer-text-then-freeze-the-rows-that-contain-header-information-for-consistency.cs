// Title: Aspose.Cells C# – Configure Header/Footer, Repeat Title Row, and Freeze Top Row
// Description: Demonstrates how to add custom left, center, and right sections to the header and footer using PageSetup, repeat the first row on every printed page, and freeze the top row (row 1) at cell A2 in an Excel workbook created with Aspose.Cells for .NET.
// Keywords: Aspose.Cells header footer C# | Aspose.Cells repeat title rows | Aspose.Cells freeze panes | PageSetup header footer Aspose | Print title rows Aspose.Cells | C# Excel header footer example | freeze first row Aspose.Cells
// Common Searches: set custom header and footer Aspose.Cells .NET | repeat header row on each printed page Aspose.Cells | freeze top row while scrolling Aspose.Cells C# | Aspose.Cells page setup header footer syntax | how to use FreezePanes in Aspose.Cells
// Developer Intent: Add a multi‑section header and footer, make row 1 repeat on printed pages, and keep row 1 visible during scrolling in an Excel file using Aspose.Cells for .NET.
// Use Cases: Generate sales or inventory reports where the file name, report title, and date appear in the header and page numbers with sheet name appear in the footer. | Print large worksheets with the column headings repeated on every page while keeping those headings frozen for on‑screen navigation.
// AI Prompts: Create C# code with Aspose.Cells that sets left, center, and right header sections, defines matching footer sections, repeats row 1 on each printed page, freezes row 1 at cell A2, and saves the workbook. | Show an Aspose.Cells example that uses PageSetup to configure a header/footer, applies PrintTitleRows = "$1:$1", and calls FreezePanes("A2", 1, 0).

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderFooterFreezeDemo
{
    // Demonstrates how to add custom left, center, and right sections to the header and footer using PageSetup, repeat the first row on every printed page, and freeze the top row (row 1) at cell A2 in an Excel workbook created with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["C1"].PutValue("Price");

            for (int i = 2; i <= 10; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue(i * 2);
                worksheet.Cells[$"C{i}"].PutValue(i * 1.5);
            }

            // Access page setup to configure header and footer
            PageSetup pageSetup = worksheet.PageSetup;

            // Header: left - file name, center - custom text, right - current date
            pageSetup.SetHeader(0, "&F");               // Left section
            pageSetup.SetHeader(1, "Sales Report");     // Center section
            pageSetup.SetHeader(2, "&D");               // Right section

            // Footer: left - page number, center - empty, right - sheet name
            pageSetup.SetFooter(0, "Page &P");          // Left section
            pageSetup.SetFooter(1, "");                 // Center section
            pageSetup.SetFooter(2, "&A");               // Right section

            // Repeat the first row on each printed page
            pageSetup.PrintTitleRows = "$1:$1";

            // Freeze the header row (row 1) so it stays visible while scrolling
            // Freeze at cell A2 with 1 frozen row and 0 frozen columns
            worksheet.FreezePanes("A2", 1, 0);

            // Save the workbook
            workbook.Save("HeaderFooterFreezeDemo.xlsx");
        }
    }
}
