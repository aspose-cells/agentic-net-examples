// Title: Repeat Row 1 as Print Title on Every Page with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds headers in row 1, fills rows 2‑100 with sample data, sets PageSetup.PrintTitleRows to "$1:$1" so the first row repeats on each printed page, defines a print area (A1:C100), and saves the file.
// Keywords: Aspose.Cells PrintTitleRows | C# repeat header row printing | Aspose.Cells PageSetup example | Excel print titles .NET | set print area Aspose.Cells
// Common Searches: Aspose.Cells repeat header row printing | C# PageSetup PrintTitleRows | How to set print titles in Aspose.Cells | Print area and title rows Aspose.Cells .NET
// Developer Intent: Configure the worksheet so that row 1 appears as a title on every printed page.
// Use Cases: Multi‑page sales reports where column headings stay visible on each sheet. | Printable invoices that need a static header row across all pages. | Exporting large data tables to Excel with a defined print area and repeated header for consistent pagination.
// AI Prompts: Generate C# code using Aspose.Cells to repeat rows 1‑2 as print titles and set a print area for columns A‑E. | Explain how to configure PageSetup to repeat column headings when printing a worksheet with Aspose.Cells. | Show how to validate that PrintTitleRows is applied before saving the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitleRowsDemo
{
    // Creates a workbook, adds headers in row 1, fills rows 2‑100 with sample data, sets PageSetup.PrintTitleRows to "$1:$1" so the first row repeats on each printed page, defines a print area (A1:C100), and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample header in the first row
            worksheet.Cells["A1"].PutValue("Header 1");
            worksheet.Cells["B1"].PutValue("Header 2");
            worksheet.Cells["C1"].PutValue("Header 3");

            // Add some data rows to demonstrate pagination
            for (int i = 2; i <= 100; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Data A{i - 1}");
                worksheet.Cells[$"B{i}"].PutValue($"Data B{i - 1}");
                worksheet.Cells[$"C{i}"].PutValue($"Data C{i - 1}");
            }

            // Access the page setup of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set the print title rows so that the first row repeats on each printed page
            // The format "$1:$1" means row 1 only.
            pageSetup.PrintTitleRows = "$1:$1";

            // Optionally define a print area that includes all data
            pageSetup.PrintArea = "A1:C100";

            // Save the workbook (lifecycle rule: save)
            workbook.Save("PrintTitleRowsDemo.xlsx");

            Console.WriteLine("Workbook saved with PrintTitleRows set to repeat row 1 on each page.");
        }
    }
}
