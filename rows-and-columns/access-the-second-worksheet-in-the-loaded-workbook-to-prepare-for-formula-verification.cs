// Title: How to access the second worksheet in an Excel workbook with Aspose.Cells (C#) and verify formula results
// AI Prompts: Load 'input.xlsx' using Aspose.Cells, select the worksheet at index 1, call workbook.CalculateFormula(), and output the evaluated value of cell B2. | Make the second worksheet the active sheet, recalculate all formulas, then save the workbook as 'output.xlsx' after confirming the cell value.
// Common Searches: Aspose.Cells C# get value of cell B2 on the second sheet after formula calculation | C# Aspose.Cells calculate formulas on a specific worksheet by index | How to set active worksheet to the second sheet in Aspose.Cells before saving | Retrieve and verify formula results on sheet index 1 using Aspose.Cells C# | Load workbook and access worksheet[1] with Aspose.Cells for formula evaluation
// Tags: access worksheet by index Aspose.Cells C# | calculate workbook formulas Aspose.Cells | retrieve cell value after formula evaluation C# | set active sheet index Aspose.Cells | save workbook after verification Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an existing Excel file, selects the second worksheet (index 1), optionally makes it the active sheet, recalculates all formulas, reads the value of cell B2 to confirm the calculation, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (zero‑based index = 1)
        Worksheet secondWorksheet = workbook.Worksheets[1];

        // Optionally make it the active sheet
        workbook.Worksheets.ActiveSheetIndex = secondWorksheet.Index;

        // Calculate all formulas in the workbook (including those on the second sheet)
        workbook.CalculateFormula();

        // Example: output the value of a cell on the second worksheet to verify calculation
        Console.WriteLine("Value of B2 on second sheet: " + secondWorksheet.Cells["B2"].Value);

        // Save the workbook after verification (if any changes were made)
        workbook.Save("output.xlsx");
    }
}
