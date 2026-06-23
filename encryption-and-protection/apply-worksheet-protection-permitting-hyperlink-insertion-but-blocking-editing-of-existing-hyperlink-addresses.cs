using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add an initial hyperlink (will be allowed to exist)
            worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

            // Access the worksheet protection settings
            Protection protection = worksheet.Protection;

            // Allow users to insert new hyperlinks while the sheet is protected
            protection.AllowInsertingHyperlink = true;

            // Disallow editing of existing cell contents (including hyperlink addresses)
            protection.AllowEditingContent = false;

            // Set a password for the protection (optional but recommended)
            protection.Password = "securePassword123";

            // Apply protection to the worksheet (protect all aspects)
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("WorksheetProtection_HyperlinkAllowed.xlsx");
        }
    }
}