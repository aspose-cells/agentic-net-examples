using Aspose.Cells;
using System;

class BatchCopyTemplateRow
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Add worksheets (the first sheet already exists)
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

        // Populate a template row (row 0) in each worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.Cells["A1"].PutValue("ID");
            ws.Cells["B1"].PutValue("Name");
            ws.Cells["C1"].PutValue("Value");
        }

        // Define the destination row indices (zero‑based) where the template row will be copied
        int[] targetRows = new int[] { 2, 5, 8 };

        // Loop through each worksheet and copy the template row to the target rows
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Cells cells = ws.Cells;

            foreach (int targetRow in targetRows)
            {
                // CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, rowNumber)
                // Here sourceCells is the same as destination cells because we copy within the sheet
                cells.CopyRows(cells, 0, targetRow, 1);
            }
        }

        // Save the workbook (lifecycle save)
        workbook.Save("BatchCopyTemplateRow.xlsx");
    }
}