using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            HideRowsAndColumnsDemo.Run();
        }
    }

    public class HideRowsAndColumnsDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (optional, just to see the effect)
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Hide rows 3 through 5 (zero‑based index 2, count 3)
            cells.HideRows(2, 3);

            // Hide columns B through D (zero‑based index 1, count 3)
            cells.HideColumns(1, 3);

            // Save the workbook to an XLSX file
            workbook.Save("HideRowsAndColumnsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}