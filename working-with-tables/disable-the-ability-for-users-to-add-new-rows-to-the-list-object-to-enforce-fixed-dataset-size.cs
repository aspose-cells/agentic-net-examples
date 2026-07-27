using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class DisableListObjectRowInsertion
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that will become the list object (table)
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");

        // Add a ListObject covering the data range (A1:B3) with headers
        int listIndex = sheet.ListObjects.Add("A1", "B3", true);
        ListObject table = sheet.ListObjects[listIndex];

        // Access the protection settings of the worksheet
        Protection protection = sheet.Protection;

        // Disable insertion and deletion of rows while the sheet is protected
        protection.AllowInsertingRow = false; // users cannot add rows
        protection.AllowDeletingRow = false;  // users cannot delete rows

        // Set a password (optional) and protect the worksheet with all protections
        protection.Password = "pwd123";
        sheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("FixedSizeTable.xlsx", SaveFormat.Xlsx);
    }
}