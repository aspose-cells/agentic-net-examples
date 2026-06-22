using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the protection settings of the worksheet
            Protection protection = worksheet.Protection;

            // Allow inserting rows while the sheet is protected
            protection.AllowInsertingRow = true;

            // Disallow deleting rows while the sheet is protected
            protection.AllowDeletingRow = false;

            // Optional: set a password for the protection
            protection.Password = "mySecretPwd";

            // Apply protection to the worksheet (protect all aspects)
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("WorksheetProtection_InsertionOnly.xlsx");
        }
    }
}