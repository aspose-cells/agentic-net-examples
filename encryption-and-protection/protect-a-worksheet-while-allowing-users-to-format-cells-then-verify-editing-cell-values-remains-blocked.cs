// Title: Protect an Excel worksheet with Aspose.Cells for .NET – allow formatting, block value edits, and verify settings
// Description: Demonstrates how to create a workbook, lock cells, enable cell‑formatting, set a password, save, reload, and confirm that content editing is disabled while formatting remains permitted.
// Keywords: Aspose.Cells worksheet protection .NET | allow cell formatting protect sheet | disable content editing Excel | worksheet password Aspose.Cells | verify IsProtected flag
// Common Searches: Aspose.Cells protect sheet but allow formatting | C# block cell value changes while permitting style changes | check worksheet protection status after saving | set password for Excel worksheet using Aspose.Cells
// Developer Intent: Apply worksheet protection that permits formatting actions but prevents any modification of cell values, then programmatically confirm the protection flags.
// Use Cases: Distribute a template where users can style cells but must not alter underlying data. | Publish a financial report that stays data‑secure yet allows conditional‑formatting tweaks. | Load a protected workbook in an automated workflow and ensure editing is disabled before processing.
// AI Prompts: Show C# code with Aspose.Cells to protect a sheet, allow only cell formatting, and set a password. | How can I programmatically verify that AllowEditingContent is false after reloading a protected workbook? | Explain how to let users change cell styles in Excel while keeping cell values read‑only using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // Demonstrates how to create a workbook, lock cells, enable cell‑formatting, set a password, save, reload, and confirm that content editing is disabled while formatting remains permitted.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put sample data into a cell (will be locked by default)
            cells["A1"].PutValue("Original Value");

            // ---------- Configure protection ----------
            // Access the protection object of the worksheet
            Protection protection = sheet.Protection;

            // Allow users to format cells but NOT edit cell contents
            protection.AllowFormattingCell = true;      // users can change formatting
            protection.AllowEditingContent = false;    // users cannot change values

            // Optionally set a password (can be null if not needed)
            protection.Password = "pwd123";

            // Apply protection to the worksheet
            sheet.Protect(ProtectionType.All);

            // ---------- Save the workbook ----------
            string filePath = "ProtectedWorksheet.xlsx";
            workbook.Save(filePath);

            // ---------- Load the workbook to verify ----------
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Protection loadedProtection = loadedSheet.Protection;

            // Verify that the worksheet is protected
            Console.WriteLine("Worksheet IsProtected: " + loadedSheet.IsProtected);

            // Verify that editing content is still disallowed
            Console.WriteLine("AllowEditingContent: " + loadedProtection.AllowEditingContent);
            // Verify that formatting cells is allowed
            Console.WriteLine("AllowFormattingCell: " + loadedProtection.AllowFormattingCell);

            // Attempt to modify a cell value programmatically
            // (Aspose.Cells allows programmatic changes regardless of UI protection,
            //  but the UI will block the edit because AllowEditingContent is false)
            try
            {
                loadedSheet.Cells["A1"].PutValue("New Value");
                Console.WriteLine("Cell value changed programmatically.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to change cell value: " + ex.Message);
            }

            // Save the workbook after the attempted edit (optional)
            loadedWorkbook.Save("ProtectedWorksheet_Verified.xlsx");
        }
    }
}
