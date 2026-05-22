using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class InsertMultiLineStringIntoTableCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add header row and a sample data row to define the table range
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Description");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Initial");

        // Create a ListObject (Excel table) that includes the header and data rows
        int tableIndex = worksheet.ListObjects.Add("A1", "B2", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Multi‑line string using newline characters
        string multiLineText = "Line 1\nLine 2\nLine 3";

        // Insert the multi‑line string into the first data row (row offset 1) and second column (column offset 1)
        table.PutCellValue(1, 1, multiLineText);

        // Save the workbook to a file
        workbook.Save("MultiLineTableCell.xlsx", SaveFormat.Xlsx);
    }
}