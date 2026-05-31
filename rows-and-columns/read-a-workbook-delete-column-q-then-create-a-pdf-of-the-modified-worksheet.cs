using System;
using Aspose.Cells;

namespace AsposeCellsColumnDeletionToPdf
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from file
            // (Replace "input.xlsx" with the path to your source workbook)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Delete column Q (zero‑based index 16)
            sheet.Cells.DeleteColumn(16);

            // Save the modified workbook as PDF
            // (Replace "output.pdf" with the desired output path)
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}