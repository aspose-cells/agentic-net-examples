// Title: Aspose.Cells C# – Retrieve Sheet2!E3 formula after deleting rows in Sheet1
// Description: Load an Excel file, delete rows 2 and 4 from the first worksheet with reference updates, then read and display the formula in cell E3 of the second worksheet. The example shows how the DeleteRow(true) flag keeps cross‑sheet formulas accurate before saving the workbook.
// Keywords: Aspose.Cells C# delete rows | DeleteRow true update references | read cell formula after row removal | Sheet2 E3 formula Aspose | cross‑sheet formula adjustment | Excel automation C# Aspose.Cells
// Common Searches: Aspose.Cells get formula after deleting rows | C# DeleteRow true effect on other sheets | How to read updated formula in another worksheet | Aspose.Cells keep formulas when rows are removed | Retrieve cell formula after structural changes
// Developer Intent: Obtain the current formula in Sheet2!E3 after removing two rows from Sheet1 while preserving reference integrity.
// Use Cases: Validate that formulas referencing deleted rows are auto‑adjusted. | Extract updated formulas for documentation or audit purposes. | Generate a report of key formulas after workbook restructuring.
// AI Prompts: Generate C# code using Aspose.Cells to delete rows 2 and 4 in Sheet1 with reference updates and then fetch the formula from Sheet2!E3. | Explain how the DeleteRow method’s ‘true’ argument updates formulas in other worksheets. | Show how to detect and handle an empty formula after rows are deleted with Aspose.Cells.

using Aspose.Cells;
using System;

// Load an Excel file, delete rows 2 and 4 from the first worksheet with reference updates, then read and display the formula in cell E3 of the second worksheet. The example shows how the DeleteRow(true) flag keeps cross‑sheet formulas accurate before saving the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Verify that the workbook has at least two worksheets
        if (workbook.Worksheets.Count < 2)
        {
            Console.WriteLine("The workbook must contain at least two worksheets.");
            return;
        }

        // ----- First deletion -----
        // Delete row 2 (zero‑based index 1) in the first worksheet.
        // The second parameter 'true' updates references in other worksheets.
        workbook.Worksheets[0].Cells.DeleteRow(1, true);

        // ----- Second deletion -----
        // Delete row 4 (zero‑based index 3) in the first worksheet.
        // After the first deletion rows have shifted, so we delete the next intended row.
        workbook.Worksheets[0].Cells.DeleteRow(3, true);

        // Retrieve the formula from cell E3 of the second worksheet (index 1)
        string formula = workbook.Worksheets[1].Cells["E3"].Formula;

        // Output the formula
        Console.WriteLine($"Formula in Sheet2!E3 after deletions: {formula}");

        // Save the workbook if you need to persist changes
        workbook.Save("output.xlsx");
    }
}
