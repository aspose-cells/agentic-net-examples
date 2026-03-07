using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------
            // Add horizontal page breaks
            // ------------------------------

            // Add a horizontal page break at row index 5 (top-left of cell A6)
            worksheet.HorizontalPageBreaks.Add(5);

            // Add a horizontal page break at row 10, spanning columns 1 to 4 (B to E)
            worksheet.HorizontalPageBreaks.Add(10, 1, 4);

            // Add a horizontal page break using a cell name (e.g., "D15")
            worksheet.HorizontalPageBreaks.Add("D15");

            // ------------------------------
            // Add vertical page breaks
            // ------------------------------

            // Add a vertical page break at column index 3 (top-left of cell D1)
            worksheet.VerticalPageBreaks.Add(3);

            // Add a vertical page break spanning rows 0 to 20 at column 6 (G)
            worksheet.VerticalPageBreaks.Add(0, 20, 6);

            // Add a vertical page break using a cell name (e.g., "H5")
            worksheet.VerticalPageBreaks.Add("H5");

            // ------------------------------
            // Save the workbook (lifecycle save)
            // ------------------------------
            workbook.Save("PageBreaksDemo.xlsx");
        }
    }
}