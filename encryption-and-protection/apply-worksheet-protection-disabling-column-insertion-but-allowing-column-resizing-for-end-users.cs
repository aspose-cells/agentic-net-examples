// Title: C# – Disable Column Insertion While Allowing Column Resizing with Aspose.Cells Worksheet Protection
// Description: Shows how to protect an Aspose.Cells worksheet, block column insertion, enable column width changes, set an optional password, and save the workbook.
// Keywords: Aspose.Cells worksheet protection | disable column insertion | allow column resizing | C# Aspose.Cells example | worksheet protection password | AllowInsertingColumn false | AllowFormattingColumn true | ProtectionType.All | Excel template protection | prevent column addition
// Common Searches: Aspose.Cells protect sheet prevent column insert | C# Aspose.Cells allow column width change only | How to set worksheet protection password in Aspose.Cells | Disable column insertion Aspose.Cells C# | Allow formatting column but not insert Aspose.Cells
// Developer Intent: Protect a worksheet so users cannot insert new columns but can resize existing ones.
// Use Cases: Distribute a financial template where column layout must stay fixed while users adjust widths for readability. | Create a shared report that blocks structural changes (column addition) yet permits column width tweaks for long text. | Secure a workbook with a password, preventing column insertion while still allowing column formatting.
// AI Prompts: Generate C# code using Aspose.Cells to protect a worksheet, forbid column insertion, allow column resizing, and set a password. | Explain how to modify the protection settings to also block row deletion while keeping column formatting enabled. | Provide a step‑by‑step guide for applying different protection options to multiple worksheets in one workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    // Shows how to protect an Aspose.Cells worksheet, block column insertion, enable column width changes, set an optional password, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the protection settings of the worksheet
            Protection protection = worksheet.Protection;

            // Disallow insertion of new columns while the sheet is protected
            protection.AllowInsertingColumn = false;

            // Allow users to resize (format) columns
            protection.AllowFormattingColumn = true;

            // Set a password for the protection (optional)
            protection.Password = "password123";

            // Apply protection to the worksheet with all protection types
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("ProtectedWorksheet.xlsx");
        }
    }
}
