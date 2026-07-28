// Title: C# Aspose.Cells Example: Verify Hidden Rows Remain Hidden After Worksheet Protection
// Description: Shows how to hide specific rows, apply ProtectionType.All to a worksheet, and confirm that the hidden state persists before protection, after protection, and after re‑loading the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | worksheet protection | hidden rows | IsRowHidden | ProtectionType.All | preserve row visibility | save and reload workbook | Excel sheet lock | Aspose.Cells API
// Common Searches: Aspose.Cells keep hidden rows after protect | Does Protect affect row visibility in Aspose.Cells | Check hidden rows after saving protected workbook .NET | IsRowHidden after worksheet protection | C# verify hidden rows in protected sheet
// Developer Intent: Confirm that applying worksheet protection does not alter the hidden status of rows or columns.
// Use Cases: Hide rows or columns, protect the sheet, and ensure the concealed elements stay hidden when the file is opened in Excel. | Automated unit tests that validate row visibility remains unchanged after protection and after persisting the workbook. | Generate reports where certain rows must stay hidden while the rest of the worksheet is locked for editing.
// AI Prompts: Create a C# snippet with Aspose.Cells that hides rows 2 and 4, protects the worksheet using ProtectionType.All, saves the workbook, reloads it, and asserts that IsRowHidden returns true for those rows. | Write an xUnit test in .NET that verifies hidden rows persist after applying sheet protection and reopening the workbook with Aspose.Cells. | Explain how worksheet protection options interact with row and column hidden properties in Aspose.Cells and whether any settings can change visibility.

using System;
using Aspose.Cells;

namespace WorksheetProtectionVisibilityDemo
{
    // Shows how to hide specific rows, apply ProtectionType.All to a worksheet, and confirm that the hidden state persists before protection, after protection, and after re‑loading the file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in rows 1 to 5 (0‑based index)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Hide rows 2 and 4 (0‑based index 1 and 3) using Row.IsHidden property
            cells.Rows[1].IsHidden = true; // Hide Row 2
            cells.Rows[3].IsHidden = true; // Hide Row 4

            // Verify hidden status before protection
            Console.WriteLine("Before protection:");
            for (int i = 0; i < 5; i++)
            {
                bool hidden = cells.IsRowHidden(i);
                Console.WriteLine($"Row {i + 1} hidden: {hidden}");
            }

            // Protect the worksheet (no password, all protection types)
            sheet.Protect(ProtectionType.All);

            // Verify hidden status after protection
            Console.WriteLine("\nAfter protection:");
            for (int i = 0; i < 5; i++)
            {
                bool hidden = cells.IsRowHidden(i);
                Console.WriteLine($"Row {i + 1} hidden: {hidden}");
            }

            // Save the workbook
            string filePath = "ProtectedVisibilityDemo.xlsx";
            workbook.Save(filePath);

            // Load the saved workbook and re‑check visibility
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Cells loadedCells = loadedSheet.Cells;

            Console.WriteLine("\nAfter loading saved file:");
            for (int i = 0; i < 5; i++)
            {
                bool hidden = loadedCells.IsRowHidden(i);
                Console.WriteLine($"Row {i + 1} hidden: {hidden}");
            }

            // Output final protection status
            Console.WriteLine($"\nWorksheet IsProtected: {loadedSheet.IsProtected}");
        }
    }
}
