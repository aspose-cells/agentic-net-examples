// Title: C# – Set Print Title Rows (Rows 1‑2) in Aspose.Cells Worksheet
// Description: Creates a new workbook, adds optional data, assigns rows 1 and 2 as repeating print titles via Worksheet.PageSetup.PrintTitleRows, and saves the file as PrintTitleRowsRows1to2.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PrintTitleRows | repeat header rows C# | Worksheet.PageSetup.PrintTitleRows example | Aspose.Cells set print titles | Excel repeat rows on each page .NET
// Common Searches: Aspose.Cells repeat rows on each printed page | C# set PrintTitleRows property | how to set print title rows in Aspose.Cells | Aspose.Cells worksheet print titles example | repeat first two rows when printing Excel with Aspose
// Developer Intent: Configure a worksheet so that rows 1‑2 appear as print titles on every printed page.
// Use Cases: Add static header rows to a multi‑page report and ensure they repeat on each printed sheet. | Automate generation of printable Excel files where column headings must appear on every page. | Create a workbook programmatically and set repeat rows without opening Excel manually.
// AI Prompts: Generate C# code that sets rows 1‑2 as print titles in an Aspose.Cells worksheet and saves the workbook. | Explain the purpose of Worksheet.PageSetup.PrintTitleRows and how to modify or clear it. | Show an example that adds dynamic data to a worksheet while keeping the first two rows as repeating print titles.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitleRowsDemo
{
    // Creates a new workbook, adds optional data, assigns rows 1 and 2 as repeating print titles via Worksheet.PageSetup.PrintTitleRows, and saves the file as PrintTitleRowsRows1to2.xlsx using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data (optional, just for illustration)
            worksheet.Cells["A1"].PutValue("Header Row 1");
            worksheet.Cells["A2"].PutValue("Header Row 2");
            for (int i = 3; i <= 20; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Data {i - 2}");
            }

            // Set rows 1 and 2 to repeat at the top of each printed page
            worksheet.PageSetup.PrintTitleRows = "$1:$2";

            // Save the workbook
            workbook.Save("PrintTitleRowsRows1to2.xlsx");
        }
    }
}
