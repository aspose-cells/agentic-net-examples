using System;
using Aspose.Cells;

class FreezeAndSaveXlsm
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["B1"].PutValue("Data1");
        sheet.Cells["C1"].PutValue("Data2");
        sheet.Cells["A2"].PutValue("Row1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(20);

        // Freeze the first column (header column)
        // Freeze at cell B1 (row index 0, column index 1) with 0 frozen rows and 1 frozen column
        sheet.FreezePanes(0, 1, 0, 1);

        // Save the workbook as an XLSM file (macro‑enabled workbook)
        workbook.Save("FrozenHeader.xlsm", SaveFormat.Xlsm);
    }
}