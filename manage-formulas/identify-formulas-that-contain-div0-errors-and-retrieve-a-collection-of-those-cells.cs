// Title: Find and list #DIV/0! cells in an Excel file with Aspose.Cells for .NET
// Description: Loads a workbook, forces formula evaluation, scans the used range of the first worksheet, checks each cell for an error, uses ErrorCellValueType.Calc to identify #DIV/0! errors, collects those cells, prints their addresses, and saves the workbook for further processing.
// Keywords: Aspose.Cells #DIV/0 error | C# detect division by zero in Excel | list error cells Aspose.Cells | Excel formula error detection .NET | retrieve cells with #DIV/0! Aspose
// Common Searches: Aspose.Cells find cells with #DIV/0! error | C# code to locate division by zero errors in Excel | how to get addresses of #DIV/0! cells using Aspose | enumerate formula errors in a workbook with Aspose.Cells
// Developer Intent: Locate every cell that contains the #DIV/0! error after formulas are calculated.
// Use Cases: Generate a data‑quality report that lists all #DIV/0! cells. | Apply conditional formatting (e.g., red fill) to highlight division‑by‑zero errors before sharing the file. | Replace #DIV/0! values with a default number to avoid downstream processing failures.
// AI Prompts: Create a C# Aspose.Cells snippet that finds #DIV/0! cells and sets their background color to red. | Show how to log row and column indices of #DIV/0! errors without scanning the entire worksheet. | Explain the mapping between Excel error types and Aspose.Cells ErrorCellValueType values and how to handle each.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads a workbook, forces formula evaluation, scans the used range of the first worksheet, checks each cell for an error, uses ErrorCellValueType.Calc to identify #DIV/0! errors, collects those cells, prints their addresses, and saves the workbook for further processing.
class IdentifyDiv0Errors
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Calculate all formulas so that error values are materialized
        workbook.CalculateFormula();

        // Access the first worksheet's cells
        Cells cells = workbook.Worksheets[0].Cells;

        // List to store cells that contain the #DIV/0! error
        List<Cell> div0ErrorCells = new List<Cell>();

        // Iterate through the used range of the worksheet
        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            for (int col = 0; col <= cells.MaxDataColumn; col++)
            {
                Cell cell = cells[row, col];

                // Check if the cell's value is an error
                if (cell.IsErrorValue)
                {
                    // Obtain the rich value to determine the specific error type
                    CellRichValue rich = cell.GetRichValue();

                    // In Aspose.Cells, #DIV/0! maps to ErrorCellValueType.Calc
                    if (rich != null && rich.ErrorValue == ErrorCellValueType.Calc)
                    {
                        div0ErrorCells.Add(cell);
                    }
                }
            }
        }

        // Output the addresses of cells that contain #DIV/0!
        Console.WriteLine("Cells containing #DIV/0! error:");
        foreach (Cell errCell in div0ErrorCells)
        {
            Console.WriteLine(errCell.Name);
        }

        // Save the workbook if further processing or inspection is required
        workbook.Save("output.xlsx");
    }
}
