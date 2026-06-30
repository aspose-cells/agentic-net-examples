using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some sample data (optional)
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].Value = $"Row {i} with some long text to test auto‑fit.";
        }

        // Determine the last row index that contains data
        int maxRow = cells.MaxDataRow; // Returns -1 if the sheet is empty

        if (maxRow >= 0)
        {
            // Auto‑fit all rows from the first row (0) to the last data row
            sheet.AutoFitRows(0, maxRow);
        }

        // Save the workbook (adjust the path as needed)
        workbook.Save("AutoFitRowsDemo.xlsx");
    }
}

// Author: Aspose.Cells .NET example – auto‑fit rows from start to max row.