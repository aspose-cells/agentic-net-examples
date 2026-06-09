using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells R2:S2 (row index 1, column index 17, 1 row, 2 columns)
        sheet.Cells.Merge(1, 17, 1, 2);

        // Protect the worksheet so that all cells (including the merged region) are read‑only
        sheet.Protect(ProtectionType.All);

        // Save the workbook as XLSX
        workbook.Save("MergedProtected.xlsx", SaveFormat.Xlsx);
    }
}