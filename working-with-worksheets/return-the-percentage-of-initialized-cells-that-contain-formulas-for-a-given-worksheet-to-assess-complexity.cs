// Title: C# – Compute Formula Density (% of Initialized Cells) in an Aspose.Cells Worksheet
// Description: Learn how to use Aspose.Cells for .NET to calculate the percentage of non‑empty cells that contain formulas in a worksheet. The example iterates the used range, counts initialized cells and formula cells, and returns the formula density, helping you gauge worksheet complexity.
// Keywords: Aspose.Cells formula density | C# calculate formula percentage | worksheet formula count .NET | initialized cells with formulas | used range cell analysis Aspose | formula complexity metric | Aspose.Cells GetFormulaPercentage
// Common Searches: how to find formula density in an Aspose.Cells worksheet | percentage of cells with formulas using C# Aspose | count non‑empty formula cells Aspose.Cells | measure worksheet complexity Aspose.Cells | C# code to calculate formula percentage in Excel
// Developer Intent: Obtain the proportion of populated cells that are formulas in a specific worksheet using Aspose.Cells for .NET.
// Use Cases: Evaluate worksheet complexity before bulk data processing. | Create a report showing formula usage across all sheets in a workbook. | Decide whether to convert formulas to static values when formula density exceeds a set threshold.
// AI Prompts: Generate a C# method with Aspose.Cells that returns the formula density (percentage of initialized cells that are formulas) for a given Worksheet. | Show an example of safely iterating the used range of a worksheet to count populated cells and formula cells, handling empty sheets gracefully. | Suggest enhancements to the GetFormulaPercentage function to ignore constant‑only cells while still counting array formulas.

using System;
using Aspose.Cells;

// Learn how to use Aspose.Cells for .NET to calculate the percentage of non‑empty cells that contain formulas in a worksheet. The example iterates the used range, counts initialized cells and formula cells, and returns the formula density, helping you gauge worksheet complexity.
class FormulaComplexityCalculator
{
    // Calculates the percentage of initialized cells that contain formulas in the given worksheet.
    static double GetFormulaPercentage(Worksheet sheet)
    {
        Cells cells = sheet.Cells;

        // Determine the used range of the worksheet.
        int maxRow = cells.MaxDataRow;      // zero‑based index of the last row with data
        int maxCol = cells.MaxDataColumn;   // zero‑based index of the last column with data

        int initializedCellCount = 0;
        int formulaCellCount = 0;

        // Iterate through the used range.
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // A cell is considered initialized if it is not empty.
                // Cell.Type == CellValueType.IsNull indicates an empty cell.
                if (cell.Type != CellValueType.IsNull)
                {
                    initializedCellCount++;

                    // Check whether the cell contains a formula.
                    if (cell.IsFormula)
                    {
                        formulaCellCount++;
                    }
                }
            }
        }

        // Avoid division by zero.
        if (initializedCellCount == 0)
            return 0.0;

        // Calculate percentage.
        return (double)formulaCellCount / initializedCellCount * 100.0;
    }

    static void Main()
    {
        // Load an existing workbook (replace with your actual file path).
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Use the first worksheet; change the index if needed.
        Worksheet worksheet = workbook.Worksheets[0];

        double formulaPercentage = GetFormulaPercentage(worksheet);

        Console.WriteLine($"Initialized cells with formulas: {formulaPercentage:F2}%");

        // Optionally, save the workbook after any modifications.
        // workbook.Save("output.xlsx");
    }
}
