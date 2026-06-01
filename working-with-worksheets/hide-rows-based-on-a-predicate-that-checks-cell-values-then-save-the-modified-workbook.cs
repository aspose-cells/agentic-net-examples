using System;
using Aspose.Cells;

class HideRowsByPredicate
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data in column A
        string[] values = { "Keep", "Hide", "Keep", "Hide", "Keep" };
        for (int i = 0; i < values.Length; i++)
        {
            // Populate cells A1, A2, ... (zero‑based row index)
            cells[i, 0].PutValue(values[i]);
        }

        // Hide rows where the cell value in column A equals "Hide"
        for (int i = 0; i < values.Length; i++)
        {
            if (cells[i, 0].StringValue == "Hide")
            {
                // HideRow uses zero‑based row index
                cells.HideRow(i);
            }
        }

        // Save the workbook with hidden rows applied
        workbook.Save("HiddenRowsExample.xlsx", SaveFormat.Xlsx);
    }
}