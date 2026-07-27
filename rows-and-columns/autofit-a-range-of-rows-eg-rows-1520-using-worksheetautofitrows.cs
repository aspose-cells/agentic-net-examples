using System;
using Aspose.Cells;

// Author: Example – auto‑fit rows 15‑20 using Aspose.Cells

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Optional: add sample data to see the effect of auto‑fit
        Cells cells = sheet.Cells;
        for (int i = 14; i <= 19; i++)
        {
            cells[i, 0].Value = $"Row {i + 1} with a long text that forces auto‑fit.";
        }

        // Auto‑fit rows 15‑20 (zero‑based indices 14‑19)
        sheet.AutoFitRows(14, 19);

        // Save the workbook (lifecycle rule)
        workbook.Save("AutoFitRows_15_20.xlsx");
    }
}