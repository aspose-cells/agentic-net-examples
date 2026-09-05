// Title: Lock specific cells, unlock others, and protect an Excel worksheet with a password using Aspose.Cells for .NET while still permitting cell formatting
// AI Prompts: Write C# code with Aspose.Cells that locks cell A1, unlocks cell B1, applies worksheet protection with a password, and keeps formatting enabled. | Show how to read the worksheet IsProtected flag and each cell's IsLocked style after protection to verify the lock status. | Demonstrate attempting to change the values of both locked and unlocked cells programmatically, then save the workbook to an .xlsx file.
// Common Searches: Aspose.Cells .NET lock a single cell while allowing other cells to be edited and keep formatting options active | how to check cell lock status after protecting a worksheet with a password in C# | programmatically modify locked cells in Aspose.Cells after worksheet protection and understand API bypass behavior
// Tags: Aspose.Cells lock specific cells worksheet protection | C# unlock cells in protected Excel sheet | verify cell lock status Aspose.Cells | allow cell formatting on protected worksheet Aspose.Cells | programmatic value change bypass UI protection Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, locks cell A1 and unlocks cell B1, applies password protection to the worksheet while allowing formatting, prints protection and lock states, attempts to modify both cells programmatically (showing API bypass of UI protection), and saves the file as ProtectedWorksheet.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "ProtectedSheet";

                // Fill some data
                sheet.Cells["A1"].PutValue("Locked Value");
                sheet.Cells["B1"].PutValue("Editable Value");

                // By default all cells are locked. Ensure A1 stays locked and B1 is unlocked.
                Style lockedStyle = workbook.CreateStyle();
                lockedStyle.IsLocked = true; // explicitly lock
                sheet.Cells["A1"].SetStyle(lockedStyle);

                Style unlockedStyle = workbook.CreateStyle();
                unlockedStyle.IsLocked = false; // unlock this cell
                sheet.Cells["B1"].SetStyle(unlockedStyle);

                // Apply protection with a password (oldPassword is not required, pass empty string)
                // Default protection options are used; specific allowances can be set via WorksheetProtection if needed.
                sheet.Protect(ProtectionType.All, "myPassword", string.Empty);

                // Verify protection settings programmatically
                bool isSheetProtected = sheet.IsProtected; // should be true
                bool isA1Locked = sheet.Cells["A1"].GetStyle().IsLocked; // true
                bool isB1Locked = sheet.Cells["B1"].GetStyle().IsLocked; // false

                Console.WriteLine($"Sheet protected: {isSheetProtected}");
                Console.WriteLine($"A1 locked (cannot edit in UI): {isA1Locked}");
                Console.WriteLine($"B1 locked (can edit in UI): {isB1Locked}");

                // Attempt to change values programmatically (will succeed because API bypasses UI protection)
                sheet.Cells["A1"].PutValue("Attempted Change");
                sheet.Cells["B1"].PutValue("Another Change");

                // Save the workbook
                string outputPath = "ProtectedWorksheet.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
