using System;
using Aspose.Cells;

namespace FreezePanesExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the first row and first column (row index 1, column index 1)
            // This keeps the header row and column visible while scrolling
            worksheet.FreezePanes(1, 1, 1, 1);

            // Save the workbook to a file
            workbook.Save("FreezePanesHeader.xlsx");
        }
    }
}