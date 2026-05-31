using System;
using Aspose.Cells;

class DeleteBlankColumnsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data with a blank column (column B will be blank)
        sheet.Cells["A1"].PutValue("Column A");
        sheet.Cells["C1"].PutValue("Column C"); // Column B is intentionally left blank
        sheet.Cells["A2"].PutValue("Data A");
        sheet.Cells["C2"].PutValue("Data C");

        // Delete all blank columns using default DeleteOptions (UpdateReference = false)
        sheet.Cells.DeleteBlankColumns();

        // Save the modified workbook
        workbook.Save("DeletedBlankColumns.xlsx", SaveFormat.Xlsx);
    }
}