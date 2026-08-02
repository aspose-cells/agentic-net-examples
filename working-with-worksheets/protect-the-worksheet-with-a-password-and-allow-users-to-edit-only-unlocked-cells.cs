// Title: C# Aspose.Cells: Password‑protect a worksheet while keeping cells B2:C3 editable
// Description: Shows how to create a Workbook, unlock the range B2:C3 by setting Style.IsLocked = false, configure protection settings (disallow edits to locked cells, prevent selection of locked cells, allow selection of unlocked cells), assign a password, and save the file as ProtectedWorksheet.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet protection | C# protect Excel sheet password | unlock cells Aspose.Cells | AllowEditingContent Aspose.Cells | AllowSelectingLockedCell | Excel template locked cells | set worksheet password .NET | protect specific range | Style.IsLocked Aspose.Cells | ProtectionType.All
// Common Searches: Aspose.Cells protect worksheet password C# | How to unlock specific cells before protecting sheet Aspose.Cells | Allow editing only certain cells in Excel using Aspose.Cells .NET | Worksheet protection options AllowEditingContent | C# code to lock sheet and keep B2:C3 editable
// Developer Intent: Add password protection to a worksheet while permitting edits only in the unlocked range B2:C3.
// Use Cases: Design data‑entry templates where only input cells are editable | Distribute financial models that protect formulas but allow users to fill summary fields | Share a report with collaborators who can modify specific cells without changing protected content | Create a read‑only dashboard with editable filter cells | Implement compliance‑driven spreadsheets that restrict changes to authorized areas
// AI Prompts: Write C# code with Aspose.Cells to protect a worksheet with a password and keep cells B2:C3 editable. | Show how to set Style.IsLocked = false for a range and configure Protection.AllowEditingContent, AllowSelectingLockedCell, and password in Aspose.Cells. | Provide an example of unlocking multiple ranges before applying ProtectionType.All in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetProtectionExample
{
    // Shows how to create a Workbook, unlock the range B2:C3 by setting Style.IsLocked = false, configure protection settings (disallow edits to locked cells, prevent selection of locked cells, allow selection of unlocked cells), assign a password, and save the file as ProtectedWorksheet.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Unlock cells B2:C3 so users can edit them after protection
                // Rows 2‑3 (index 1‑2) and columns B‑C (index 1‑2)
                for (int row = 1; row <= 2; row++)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        // Retrieve the cell style, modify the lock flag, and reapply
                        Style style = sheet.Cells[row, col].GetStyle();
                        style.IsLocked = false;               // Correct property name
                        sheet.Cells[row, col].SetStyle(style);
                    }
                }

                // Configure worksheet protection options
                Protection protection = sheet.Protection;
                protection.AllowEditingContent = false;          // Disallow editing locked cells
                protection.AllowSelectingLockedCell = false;     // Prevent selecting locked cells
                protection.AllowSelectingUnlockedCell = true;    // Allow selecting unlocked cells
                protection.Password = "SecurePass123";           // Set worksheet password

                // Apply protection (all protection types)
                sheet.Protect(ProtectionType.All);

                // Define output file path
                string outputPath = "ProtectedWorksheet.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
