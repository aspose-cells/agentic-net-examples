using System;
using Aspose.Cells;

class AutoFitAllRows
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (optional, demonstrates the effect)
        sheet.Cells["A1"].PutValue("First row with a long text that should cause auto‑fit.");
        sheet.Cells["A2"].PutValue("Second row");
        sheet.Cells["A3"].PutValue("Third row with even longer text that will wrap and increase row height.");

        // Determine the last row that contains data (zero‑based index)
        int maxRow = sheet.Cells.MaxDataRow;

        // Auto‑fit all rows from the first row (0) to the last data row
        sheet.AutoFitRows(0, maxRow);

        // Save the workbook
        workbook.Save("AutoFitAllRows.xlsx");
    }
}