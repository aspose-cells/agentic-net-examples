using Aspose.Cells;
using System;

class SetColumnCWidth
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set column C (zero‑based index 2) width to 20 character units
        worksheet.Cells.SetColumnWidth(2, 20);

        // Example data to illustrate the width
        worksheet.Cells["C1"].PutValue("This is a longer text entry that fits the column width.");

        // Save the workbook
        workbook.Save("ColumnCWidth.xlsx");
    }
}