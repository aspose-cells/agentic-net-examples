using System;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Access the protection settings of the worksheet
            Protection protection = sheet.Protection;

            // Disallow inserting new columns while the sheet is protected
            protection.AllowInsertingColumn = false;

            // Allow users to resize (format) columns
            protection.AllowFormattingColumn = true;

            // Optional: set a password for the protection
            protection.Password = "securePassword";

            // Apply protection to the worksheet (protect all aspects)
            sheet.Protect(ProtectionType.All);

            // Save the workbook to a file
            workbook.Save("WorksheetProtection.xlsx");
        }
    }
}