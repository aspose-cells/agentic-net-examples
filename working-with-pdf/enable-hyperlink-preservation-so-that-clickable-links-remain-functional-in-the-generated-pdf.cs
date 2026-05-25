using System;
using Aspose.Cells;

namespace HyperlinkPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put display text into a cell
            sheet.Cells["A1"].PutValue("Visit Aspose");

            // Add a hyperlink to the cell (A1)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, address
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.aspose.com");

            // Save the workbook as PDF (lifecycle: save)
            // Hyperlinks are preserved automatically in PDF output
            workbook.Save("HyperlinkPreserved.pdf", SaveFormat.Pdf);
        }
    }
}