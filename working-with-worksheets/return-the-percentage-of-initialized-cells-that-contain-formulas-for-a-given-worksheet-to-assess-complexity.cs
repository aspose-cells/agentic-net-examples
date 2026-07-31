// Title: Compute Formula Cell Percentage in an Aspose.Cells Worksheet (C#)
// Description: Loads an Excel file with Aspose.Cells, scans the used range of a worksheet, counts cells that contain a value or a formula, determines how many of those are formulas, and calculates the percentage of initialized cells that are formulas.
// Keywords: Aspose.Cells | C# formula cell percentage | count initialized cells | worksheet used range | cell.IsFormula | Excel formula density | Aspose.Cells performance | calculate formula ratio | Excel cell statistics | Aspose.Cells .NET
// Common Searches: Aspose.Cells calculate percentage of formula cells | C# count cells with formulas in Excel using Aspose | How to get formula density in a worksheet Aspose.Cells | Determine initialized cells vs formula cells Aspose | Aspose.Cells used range cell count
// Developer Intent: Identify the proportion of non‑empty cells that are formulas within a worksheet.
// Use Cases: Measure formula density to evaluate worksheet complexity before performance tuning. | Create a summary report that shows formula usage across all sheets in a workbook. | Validate that a worksheet meets a maximum allowed formula‑to‑value ratio prior to publishing.
// AI Prompts: Generate C# code using Aspose.Cells that returns the formula‑cell percentage for each worksheet in a workbook. | Explain how to skip completely empty rows and columns to speed up the formula‑percentage calculation. | Suggest a safe way to handle worksheets with zero initialized cells to avoid division‑by‑zero errors while reporting the ratio.

using System;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, scans the used range of a worksheet, counts cells that contain a value or a formula, determines how many of those are formulas, and calculates the percentage of initialized cells that are formulas.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string filePath = "input.xlsx";
        Workbook workbook = new Workbook(filePath);

        // Get the first worksheet (or any specific worksheet you need)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        int totalInitialized = 0;   // Cells that contain a value or a formula
        int formulaCount = 0;       // Cells that contain a formula

        // Iterate through the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // A cell is considered initialized if it has a value or a formula
                if (cell.Value != null || cell.IsFormula)
                {
                    totalInitialized++;

                    if (cell.IsFormula)
                    {
                        formulaCount++;
                    }
                }
            }
        }

        // Calculate the percentage of initialized cells that contain formulas
        double percentage = totalInitialized == 0
            ? 0
            : (double)formulaCount / totalInitialized * 100.0;

        // Output the results
        Console.WriteLine($"Initialized cells: {totalInitialized}");
        Console.WriteLine($"Formula cells: {formulaCount}");
        Console.WriteLine($"Percentage of formula cells: {percentage:F2}%");
    }
}
