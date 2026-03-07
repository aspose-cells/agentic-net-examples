using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill some data so that page breaks can be visualized
        for (int i = 0; i < 50; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }
        for (int j = 0; j < 20; j++)
        {
            sheet.Cells[0, j].PutValue($"Col {j + 1}");
        }

        // ---------- Horizontal page breaks ----------
        // Simple row index (break after row 5)
        sheet.HorizontalPageBreaks.Add(5);

        // Row and column (break after row 15, starting at column 0)
        sheet.HorizontalPageBreaks.Add(15, 0);

        // Row with start and end columns (break after row 25, columns 2‑6)
        sheet.HorizontalPageBreaks.Add(25, 2, 6);

        // Using cell name (break at cell G30)
        sheet.HorizontalPageBreaks.Add("G30");

        // Using Worksheet.AddPageBreaks (adds a horizontal break)
        sheet.AddPageBreaks("B2");
        sheet.AddPageBreaks("B12");

        // ---------- Vertical page breaks ----------
        // Simple column index (break after column 3)
        sheet.VerticalPageBreaks.Add(3);

        // StartRow, EndRow, Column (break spanning rows 0‑10 at column 4)
        sheet.VerticalPageBreaks.Add(0, 10, 4);

        // Row and column (break after column 5 at row 20)
        sheet.VerticalPageBreaks.Add(20, 5);

        // Using cell name (break at cell H5)
        sheet.VerticalPageBreaks.Add("H5");

        // Save the workbook with the page breaks applied
        workbook.Save("PageBreaksDemo.xlsx");
    }
}