// Title: Worksheet protection preserves hidden rows and columns – Aspose.Cells for .NET example
// Description: Demonstrates how to hide rows (and optionally columns), protect a worksheet with Aspose.Cells, and verify that the hidden state remains unchanged before and after protection, then saves the workbook.
// Keywords: Aspose.Cells protect worksheet | hidden rows after protection | IsRowHidden Aspose.Cells | IsColumnHidden after Protect | .NET spreadsheet security | worksheet visibility protection | Aspose.Cells hide rows example
// Common Searches: Does protecting a worksheet hide rows in Aspose.Cells? | Aspose.Cells keep hidden rows after sheet protection | Check column visibility after Worksheet.Protect | IsRowHidden returns false after Protect? | How to protect sheet without revealing hidden rows .NET
// Developer Intent: Verify that applying Worksheet.Protect does not modify the hidden status of rows or columns in the same sheet.
// Use Cases: Automated validation that hidden rows stay hidden when a workbook is distributed. | Creating reports with confidential rows hidden, then securing the sheet for end‑users. | Ensuring column hiding behaves identically to row hiding after applying protection.
// AI Prompts: Generate a unit test using Aspose.Cells that asserts hidden rows remain hidden after Worksheet.Protect. | Show code to hide specific columns, protect the worksheet with a password, and confirm column visibility with IsColumnHidden. | Provide a script that logs hidden status of rows and columns before and after protection, then saves the file.

using System;
using Aspose.Cells;

namespace WorksheetProtectionVisibilityDemo
{
    // Demonstrates how to hide rows (and optionally columns), protect a worksheet with Aspose.Cells, and verify that the hidden state remains unchanged before and after protection, then saves the workbook.
    public class Run
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Hide specific rows (0‑based index)
            cells.HideRow(2); // Hide Row 3
            cells.HideRow(5); // Hide Row 6

            // Verify hidden status before protection
            Console.WriteLine("Before protection:");
            Console.WriteLine($"Row 3 hidden: {cells.IsRowHidden(2)}");
            Console.WriteLine($"Row 6 hidden: {cells.IsRowHidden(5)}");

            // Protect the worksheet (no password, all protection types)
            sheet.Protect(ProtectionType.All);

            // Verify hidden status after protection
            Console.WriteLine("\nAfter protection:");
            Console.WriteLine($"Row 3 hidden: {cells.IsRowHidden(2)}");
            Console.WriteLine($"Row 6 hidden: {cells.IsRowHidden(5)}");

            // Save the workbook (optional, just to complete lifecycle)
            workbook.Save("WorksheetProtectionVisibilityDemo.xlsx");
        }
    }
}
