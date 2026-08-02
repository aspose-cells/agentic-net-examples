using System;
using Aspose.Cells;

class HideColumnsInOds
{
    static void Main()
    {
        // Load the source ODS workbook
        string sourcePath = "input.ods";
        Workbook workbook = new Workbook(sourcePath);

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide specific columns (zero‑based indexes)
        // Example: hide column B (index 1)
        cells.HideColumn(1);
        // Example: hide columns D, E, and F (start at index 3, hide 3 columns)
        cells.HideColumns(3, 3);

        // Save the modified workbook as XLSX
        string destPath = "output.xlsx";
        workbook.Save(destPath, SaveFormat.Xlsx);
    }
}