using System;
using Aspose.Cells;

class ReadFormulaAfterDeletions
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string filePath = "input.xlsx";
        Workbook workbook = new Workbook(filePath);

        // Ensure there are at least two worksheets
        if (workbook.Worksheets.Count < 2)
        {
            Console.WriteLine("The workbook must contain at least two worksheets.");
            return;
        }

        // Reference to the first worksheet (where deletions will be performed)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // First deletion: delete row 2 (zero‑based index 1)
        cells.DeleteRow(1, true);

        // Second deletion: delete row 4 (original index 3, after first deletion it becomes 2)
        cells.DeleteRow(2, true);

        // After deletions, read the formula from cell E3 of the second worksheet (index 1)
        Worksheet secondSheet = workbook.Worksheets[1];
        string formula = secondSheet.Cells["E3"].Formula;

        // Output the retrieved formula
        Console.WriteLine($"Formula in Sheet2!E3 after deletions: {formula}");

        // Optionally save the modified workbook
        workbook.Save("output.xlsx");
    }
}