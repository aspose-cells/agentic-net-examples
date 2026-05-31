using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data and formulas (optional)
        for (int i = 0; i < 20; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
            worksheet.Cells[i, 1].Formula = $"=SUM(A{i + 1}:A{i + 2})";
        }

        // Hide rows 5 through 15 (zero‑based index 4, total 11 rows)
        worksheet.Cells.HideRows(4, 11);

        // Enable formula view on the worksheet
        worksheet.ShowFormulas = true;

        // Save the workbook
        workbook.Save("HiddenRowsShowFormulas.xlsx");
    }
}