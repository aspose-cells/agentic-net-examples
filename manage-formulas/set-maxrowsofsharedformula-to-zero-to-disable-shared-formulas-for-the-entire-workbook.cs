// Title: Aspose.Cells for .NET – Disable Shared Formulas for an Entire Workbook (MaxRowsOfSharedFormula = 0)
// Description: Creates a new Workbook, sets workbook.Settings.MaxRowsOfSharedFormula to 0 to turn off shared formulas for every sheet, and saves the file as Workbook_NoSharedFormulas.xlsx. This setting forces each formula to be stored individually.
// Keywords: Aspose.Cells disable shared formulas | MaxRowsOfSharedFormula zero | turn off shared formulas .NET | Workbook.Settings.MaxRowsOfSharedFormula | Aspose.Cells workbook configuration
// Common Searches: how to disable shared formulas in Aspose.Cells .NET | set MaxRowsOfSharedFormula to 0 example | Aspose.Cells workbook settings for formulas | remove shared formulas from Excel file using Aspose | disable shared formula limit Aspose.Cells
// Developer Intent: Turn off shared formulas for the whole workbook.
// Use Cases: Generate Excel files compatible with legacy versions that do not support shared formulas. | Create workbooks where each formula is independent, simplifying downstream calculations. | Export data with explicit formulas to avoid shared‑formula size limits in large spreadsheets.
// AI Prompts: Show a C# snippet that disables shared formulas for an entire workbook using Aspose.Cells. | Explain what happens to file size and calculation performance when MaxRowsOfSharedFormula is set to zero. | Provide step‑by‑step instructions to turn off shared formulas in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace WorkbookSettingsExample
{
    // Creates a new Workbook, sets workbook.Settings.MaxRowsOfSharedFormula to 0 to turn off shared formulas for every sheet, and saves the file as Workbook_NoSharedFormulas.xlsx. This setting forces each formula to be stored individually.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Disable shared formulas for the entire workbook by setting the limit to zero
            workbook.Settings.MaxRowsOfSharedFormula = 0;

            // Save the workbook to a file (adjust the path as needed)
            workbook.Save("Workbook_NoSharedFormulas.xlsx");
        }
    }
}
