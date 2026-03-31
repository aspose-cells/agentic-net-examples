using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (50 rows, 10 columns)
            for (int row = 0; row < 50; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Insert horizontal page breaks to control row pagination
            // Break after row 10 (zero‑based index 9) and after row 30 (index 29)
            worksheet.HorizontalPageBreaks.Add(9);   // Page break after row 10
            worksheet.HorizontalPageBreaks.Add(29);  // Page break after row 30

            // Insert a vertical page break to control column pagination
            // Break after column 5 (zero‑based index 4)
            worksheet.VerticalPageBreaks.Add(4);     // Page break after column 5

            // Save the workbook; the page breaks will affect pagination when exported/printed
            workbook.Save("PageBreaksDemo.xlsx");
        }
    }
}