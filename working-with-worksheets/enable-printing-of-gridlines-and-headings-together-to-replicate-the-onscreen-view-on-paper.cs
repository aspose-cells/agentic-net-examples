using System;
using Aspose.Cells;

namespace AsposeCellsPrintGridlinesAndHeadings
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data so the printed page has content
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(2.5);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(1.8);

            // Ensure gridlines are visible on the sheet (optional, affects on‑screen view)
            worksheet.IsGridlinesVisible = true;

            // Enable printing of gridlines
            worksheet.PageSetup.PrintGridlines = true;

            // Enable printing of row and column headings
            worksheet.PageSetup.PrintHeadings = true;

            // Define the print area to include the data (optional, but keeps output tidy)
            worksheet.PageSetup.PrintArea = "A1:B3";

            // Save the workbook (save rule)
            workbook.Save("PrintGridlinesAndHeadings.xlsx");
        }
    }
}