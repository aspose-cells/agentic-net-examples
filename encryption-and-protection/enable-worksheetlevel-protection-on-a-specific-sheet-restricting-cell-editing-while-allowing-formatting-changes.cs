// Title: Protect a worksheet in Aspose.Cells for .NET – allow formatting while blocking content edits
// Description: Demonstrates how to enable worksheet‑level protection in Aspose.Cells (C#) by disabling content changes (AllowEditingContent = false), permitting cell formatting (AllowFormattingCell = true), applying full protection with sheet.Protect(ProtectionType.All), and saving the workbook.
// Keywords: Aspose.Cells worksheet protection | C# protect sheet allow formatting | AllowEditingContent false | AllowFormattingCell true | protect specific worksheet .NET | Excel sheet protection Aspose | read‑only template Aspose.Cells
// Common Searches: Aspose.Cells protect worksheet but allow formatting | C# worksheet protection AllowEditingContent false | How to lock cells content in Aspose.Cells | Enable sheet protection with formatting options Aspose | Protect only one sheet in Aspose.Cells workbook
// Developer Intent: Apply worksheet protection that prevents users from editing cell values while still allowing them to change cell styles and formats on a chosen sheet.
// Use Cases: Distribute a financial template where formulas stay locked but users can adjust colors, borders, and column widths. | Create a read‑only report that lets reviewers highlight cells without altering underlying data. | Provide a data‑entry form where only visual formatting is editable to maintain data integrity.
// AI Prompts: Generate C# code with Aspose.Cells that protects a worksheet, disables content editing, and enables cell formatting. | Show how to set AllowEditingContent to false and AllowFormattingCell to true on a specific sheet using Aspose.Cells. | Explain additional protection options (e.g., AllowInsertingRows, AllowDeletingColumns) after applying worksheet protection in Aspose.Cells.

using System;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    // Demonstrates how to enable worksheet‑level protection in Aspose.Cells (C#) by disabling content changes (AllowEditingContent = false), permitting cell formatting (AllowFormattingCell = true), applying full protection with sheet.Protect(ProtectionType.All), and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (or any specific sheet by index/name)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the protection settings for the worksheet
            Protection protection = sheet.Protection;

            // Disallow editing of locked cells' content
            protection.AllowEditingContent = false;

            // Allow users to format cells even when the sheet is protected
            protection.AllowFormattingCell = true;

            // Apply protection to the worksheet (all protection types)
            sheet.Protect(ProtectionType.All);

            // Save the workbook with the protected worksheet
            workbook.Save("ProtectedSheet.xlsx");
        }
    }
}
