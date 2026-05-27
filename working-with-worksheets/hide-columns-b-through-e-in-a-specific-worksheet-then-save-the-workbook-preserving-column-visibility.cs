using System;
using Aspose.Cells;

class HideColumnsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (you can also use workbook.Worksheets["SheetName"])
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide columns B through E (zero‑based indices 1 to 4)
        int startColumn = 1;      // Column B
        int columnCount = 4;      // B, C, D, E
        worksheet.Cells.HideColumns(startColumn, columnCount);

        // Optional: add some data to visible columns for demonstration
        worksheet.Cells["A1"].PutValue("Visible Column A");
        worksheet.Cells["F1"].PutValue("Visible Column F");

        // Save the workbook; hidden columns are retained in the file
        workbook.Save("HiddenColumns_BtoE.xlsx", SaveFormat.Xlsx);
    }
}