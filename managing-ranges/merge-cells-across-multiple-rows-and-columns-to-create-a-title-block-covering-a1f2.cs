using System;
using Aspose.Cells;

namespace MergeTitleBlockDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells from A1 (row 0, column 0) to F2 (row 1, column 5)
            // totalRows = 2 (rows 0 and 1), totalColumns = 6 (columns 0 through 5)
            cells.Merge(0, 0, 2, 6);

            // Set the title text in the merged cell (upper‑left cell of the range)
            cells[0, 0].PutValue("Report Title");

            // Apply center alignment and bold font to the title
            Style style = cells[0, 0].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.Font.IsBold = true;
            cells[0, 0].SetStyle(style);

            // Save the workbook to a file
            workbook.Save("TitleBlock.xlsx");
        }
    }
}