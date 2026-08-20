// Title: C# – Delete columns B‑D in an Aspose.Cells worksheet and auto‑update formulas
// Description: Creates a workbook, populates columns A‑E, adds a SUM formula referencing B1:D1, then removes columns B through D using DeleteOptions.UpdateReference so the formula adjusts automatically, and saves the result.
// Keywords: Aspose.Cells DeleteColumns C# | Delete multiple columns .NET | UpdateReference option | Preserve formulas after column removal | C# spreadsheet column deletion | Aspose.Cells worksheet restructuring | Aspose.Cells US developers | Aspose.Cells Europe examples
// Common Searches: Aspose.Cells delete columns B to D C# | How to keep formulas when deleting columns in Aspose.Cells | DeleteColumns with UpdateReference example | Remove range of columns without breaking formulas Aspose.Cells | C# delete columns and adjust references
// Developer Intent: Remove columns B‑D from a worksheet while automatically updating any formulas that referenced those columns.
// Use Cases: Clean up a generated report by deleting placeholder columns and ensuring summary formulas still calculate correctly. | Trim imported data to the required fields before exporting, preserving dependent calculations such as totals or averages. | Programmatically restructure a spreadsheet layout by removing unnecessary columns without breaking existing formulas.
// AI Prompts: Show me C# code that deletes columns 2 through 4 in an Aspose.Cells worksheet and updates all related formulas. | Explain how DeleteOptions.UpdateReference works when deleting a range of columns and what happens to formulas that referenced the deleted range. | Generate a complete example that deletes columns B‑D, verifies the formula adjustment, and saves the workbook.

using System;
using Aspose.Cells;

// Creates a workbook, populates columns A‑E, adds a SUM formula referencing B1:D1, then removes columns B through D using DeleteOptions.UpdateReference so the formula adjusts automatically, and saves the result.
class DeleteColumnsExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in columns A to E
        for (int col = 0; col < 5; col++)
        {
            cells[0, col].PutValue($"Header {(char)('A' + col)}");
            cells[1, col].PutValue(col + 1);
        }

        // Add a formula that references columns B through D
        cells["F1"].Formula = "=SUM(B1:D1)";

        // Configure delete options to update references after deletion
        DeleteOptions options = new DeleteOptions { UpdateReference = true };

        // Delete columns B (index 1) through D (index 3) – total 3 columns
        cells.DeleteColumns(1, 3, options);

        // Save the modified workbook
        workbook.Save("DeletedColumns.xlsx");
    }
}
