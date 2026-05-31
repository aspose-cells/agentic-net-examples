using System;
using Aspose.Cells;

namespace HideRowsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide rows 20 through 30 (zero‑based index: start at 19, hide 11 rows)
            worksheet.Cells.HideRows(19, 11);

            // Save the workbook with the rows concealed
            workbook.Save("RowsHidden.xlsx");
        }
    }
}