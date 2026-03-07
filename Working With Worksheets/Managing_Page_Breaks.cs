using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data to visualize page breaks
            for (int row = 0; row < 50; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ---------- Horizontal Page Breaks ----------
            // Add a simple horizontal page break after row 9 (zero‑based index)
            worksheet.HorizontalPageBreaks.Add(9);

            // Add a horizontal page break that spans columns 2 to 6 on row 19
            worksheet.HorizontalPageBreaks.Add(19, 2, 6);

            // Add a horizontal page break at a specific cell (D15)
            worksheet.HorizontalPageBreaks.Add("D15");

            // Use Worksheet.AddPageBreaks method to add another break at B30
            worksheet.AddPageBreaks("B30");

            // ---------- Vertical Page Breaks ----------
            // Add a simple vertical page break after column 4
            worksheet.VerticalPageBreaks.Add(4);

            // Add a vertical page break that spans rows 0 to 30 at column 7
            worksheet.VerticalPageBreaks.Add(0, 30, 7);

            // Add a vertical page break at a specific cell (G5)
            worksheet.VerticalPageBreaks.Add("G5");

            // Save the workbook to verify the page breaks
            workbook.Save("PageBreaksDemo.xlsx");
        }
    }
}