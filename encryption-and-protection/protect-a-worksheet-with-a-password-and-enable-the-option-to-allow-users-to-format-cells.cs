// Title: C# – Protect an Excel worksheet with a password while allowing cell formatting using Aspose.Cells
// Description: Shows how to set AllowFormattingCell, apply password protection (ProtectionType.All) to a worksheet, and save the workbook as an .xlsx file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | worksheet protection | password protection | AllowFormattingCell | protect sheet programmatically | cell formatting on protected sheet | ProtectionType.All | Excel security | Aspose.Cells .NET example
// Common Searches: Aspose.Cells protect worksheet password C# | allow cell formatting on a protected Excel sheet Aspose | set AllowFormattingCell before sheet.Protect | C# code to protect Excel sheet but keep formatting enabled | Aspose.Cells worksheet protection options
// Developer Intent: Apply password protection to a worksheet while keeping cell‑formatting features available to end users.
// Use Cases: Create a new workbook, protect the first sheet with a password, and let users change cell styles without unprotecting the sheet. | Add password security to an existing worksheet while selectively enabling actions such as formatting cells, inserting rows, or editing objects. | Generate a distributable Excel file that enforces data integrity yet permits visual formatting adjustments by recipients.
// AI Prompts: Generate C# code that protects an Excel worksheet with a password and enables only cell formatting using Aspose.Cells. | Explain how to combine multiple Allow* properties (e.g., AllowFormattingCell, AllowInsertingRows) before calling sheet.Protect. | Show the steps to change the password of a protected worksheet while preserving previously set AllowFormattingCell settings.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to set AllowFormattingCell, apply password protection (ProtectionType.All) to a worksheet, and save the workbook as an .xlsx file with Aspose.Cells for .NET.
    public class ProtectWorksheetAllowFormattingCell
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Get the protection object for the worksheet
                Protection protection = sheet.Protection;

                // Allow users to format cells even when the sheet is protected
                protection.AllowFormattingCell = true;

                // Protect the worksheet with a password and enable all protection types
                // The third parameter (oldPassword) is null because the sheet is not previously protected
                sheet.Protect(ProtectionType.All, "MySecretPassword", null);

                // Save the workbook to a file
                workbook.Save("ProtectedWorksheet.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectWorksheetAllowFormattingCell.Run();
        }
    }
}
