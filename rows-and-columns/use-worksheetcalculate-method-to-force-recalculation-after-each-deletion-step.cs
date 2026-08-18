// Title: C# – Recalculate formulas after row deletions with Aspose.Cells Worksheet.Calculate
// Description: Creates a workbook, fills column A with values 1‑10, adds a SUM formula in B1, then deletes rows (single, multiple, and without reference updates). After each DeleteRow/DeleteRows call, Workbook.CalculateFormula() is invoked to refresh dependent formulas before saving the file.
// Keywords: Aspose.Cells | Worksheet.Calculate | CalculateFormula | row deletion | C# example | update formulas after delete | DeleteRow overload | DeleteRows overload | Excel automation | recalculate formulas
// Common Searches: Aspose.Cells recalculate formulas after deleting rows C# | Worksheet.Calculate after DeleteRow | How to refresh SUM formula after row removal in Aspose.Cells | DeleteRow updateReference true false example | C# Aspose.Cells delete rows and recalc workbook
// Developer Intent: Refresh dependent formulas immediately after removing rows in an Aspose.Cells workbook.
// Use Cases: Delete a specific row (e.g., row 3) and instantly update a SUM formula that references the column. | Remove a block of consecutive rows and ensure aggregate functions like AVERAGE or COUNT reflect the new range. | Delete a row without adjusting references, then manually trigger recalculation to keep original formula links.
// AI Prompts: Show C# code that calls Workbook.CalculateFormula after each DeleteRow/DeleteRows operation in Aspose.Cells. | Provide an Aspose.Cells example demonstrating both DeleteRow overloads and the required Worksheet.Calculate call. | Explain when to use DeleteRow with updateReference true versus false and how Worksheet.Calculate impacts formula results.

using System;
using Aspose.Cells;

namespace AsposeCellsRowDeletionRecalc
{
    // Creates a workbook, fills column A with values 1‑10, adds a SUM formula in B1, then deletes rows (single, multiple, and without reference updates). After each DeleteRow/DeleteRows call, Workbook.CalculateFormula() is invoked to refresh dependent formulas before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data and formulas
                // Values in column A
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue(i + 1); // A1..A10 = 1..10
                }

                // Formula in B1 that sums the first 10 values of column A
                cells[0, 1].Formula = "=SUM(A1:A10)";

                // Initial calculation to set formula result
                workbook.CalculateFormula(); // Force calculation after initial data entry

                // ----- Deletion Step 1 -----
                // Delete the third row (index 2) and update references
                cells.DeleteRow(2, true);

                // Recalculate worksheet after deletion
                workbook.CalculateFormula(); // Forces recalculation of dependent formulas

                // ----- Deletion Step 2 -----
                // Delete two consecutive rows starting from row 5 (original index 4)
                // After the previous deletion, rows have shifted, so we delete at index 3
                cells.DeleteRows(3, 2, true);

                // Recalculate worksheet after second deletion
                workbook.CalculateFormula();

                // ----- Deletion Step 3 -----
                // Delete the first row without updating references (demonstrates both overloads)
                cells.DeleteRow(0, false);

                // Recalculate worksheet after third deletion
                workbook.CalculateFormula();

                // Save the modified workbook (lifecycle save)
                workbook.Save("RowDeletionWithRecalc.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
