using System;
using Aspose.Cells;

class SharedWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Enable shared mode so multiple users can edit concurrently
        workbook.Settings.Shared = true;

        // Optional: protect the shared workbook with a password
        workbook.ProtectSharedWorkbook("Password123");

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");

        // Save the shared workbook to disk
        workbook.Save("SharedWorkbook.xlsx");

        // Clean up resources
        workbook.Dispose();
    }
}