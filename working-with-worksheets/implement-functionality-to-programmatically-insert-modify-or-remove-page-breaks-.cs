using System;
using Aspose.Cells;

namespace PageBreakDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (10 rows x 5 columns)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // -------------------------------------------------
            // Insert horizontal page breaks
            // -------------------------------------------------
            // Add a break after row 3 (zero‑based index 2)
            sheet.HorizontalPageBreaks.Add(2);
            // Add a break at cell "B7" (row 6, column 1)
            sheet.HorizontalPageBreaks.Add("B7");
            // Add a break spanning columns 1 to 3 on row 9
            sheet.HorizontalPageBreaks.Add(8, 1, 3);

            // -------------------------------------------------
            // Insert vertical page breaks
            // -------------------------------------------------
            // Add a break after column 2 (zero‑based index 1)
            sheet.VerticalPageBreaks.Add(1);
            // Add a break at cell "D4" (row 3, column 3)
            sheet.VerticalPageBreaks.Add("D4");
            // Add a break spanning rows 0 to 5 on column 4
            sheet.VerticalPageBreaks.Add(0, 5, 4);

            // -------------------------------------------------
            // Modify page breaks
            // -------------------------------------------------
            // Change the first horizontal page break to row 5
            if (sheet.HorizontalPageBreaks.Count > 0)
            {
                // Remove the existing break
                sheet.HorizontalPageBreaks.RemoveAt(0);
                // Insert a new break at the desired row
                sheet.HorizontalPageBreaks.Add(4); // row index 4 => after row 5
            }

            // -------------------------------------------------
            // Remove a specific vertical page break
            // -------------------------------------------------
            // Remove the vertical page break we added at column 1 (index 0)
            if (sheet.VerticalPageBreaks.Count > 0)
            {
                sheet.VerticalPageBreaks.RemoveAt(0);
            }

            // -------------------------------------------------
            // Save the workbook to a file
            // -------------------------------------------------
            workbook.Save("PageBreakDemo.xlsx");
        }
    }
}