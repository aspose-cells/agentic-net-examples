using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the worksheet protection settings
            Protection protection = worksheet.Protection;

            // Allow users to select locked cells only
            protection.AllowSelectingLockedCell = true;
            protection.AllowSelectingUnlockedCell = false; // optional, default is false

            // Protect the worksheet with password "Secure123"
            // Using the overload that accepts protection type, password and old password (null because it's a new protection)
            worksheet.Protect(ProtectionType.All, "Secure123", null);

            // Save the workbook
            workbook.Save("ProtectedWorksheet.xlsx");
        }
    }
}