using System;
using Aspose.Cells;

namespace ProtectWorksheetExample
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or any specific worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Configure protection settings to prevent row and column deletion
            Protection protection = sheet.Protection;
            protection.AllowDeletingRow = false;      // Disallow deleting rows
            protection.AllowDeletingColumn = false;   // Disallow deleting columns

            // Optionally set a password (can be null or empty if not needed)
            protection.Password = "securePassword";

            // Apply protection to the worksheet (protect all aspects)
            sheet.Protect(ProtectionType.All, protection.Password, null);

            // Save the protected workbook
            workbook.Save("output.xlsx");
        }
    }
}