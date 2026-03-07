using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HideRowsAndColumnsDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (10 rows x 10 columns)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Hide rows 3 to 5 (zero‑based index 2, total 3 rows)
            cells.HideRows(2, 3);

            // Hide columns C to E (zero‑based index 2, total 3 columns)
            cells.HideColumns(2, 3);

            // Save the workbook
            workbook.Save("HideRowsColumnsDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            HideRowsAndColumnsDemo.Run();
        }
    }
}