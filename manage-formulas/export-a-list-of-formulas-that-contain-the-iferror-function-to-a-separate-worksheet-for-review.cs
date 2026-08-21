// Title: C# – Export IFERROR Formulas to a Review Worksheet with Aspose.Cells
// Description: Loads an Excel workbook, scans a worksheet for cells whose formulas contain the IFERROR function, records each cell address and formula, creates a new sheet named "IFERROR Review", writes the collected data with headers, and saves the updated file.
// Keywords: Aspose.Cells | C# | IFERROR extraction | export formulas | Excel automation | list cell formulas | review worksheet | error‑handling formulas | cell address lookup | Excel .NET library
// Common Searches: Aspose.Cells find IFERROR formulas C# | export cells with IFERROR to another sheet | list IFERROR formula addresses using .NET | create review worksheet for error handling formulas | scan workbook for specific function Aspose.Cells
// Developer Intent: Collect every IFERROR formula from a source sheet and write its address and expression to a new worksheet for review.
// Use Cases: Audit a workbook to locate all error‑handling formulas before publishing. | Generate a documentation report that shows where IFERROR is used and its exact syntax. | Identify formulas that may require optimization or replacement by summarizing them on a separate sheet. | Prepare a training example that demonstrates error handling across the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that extracts all IFERROR formulas and writes their addresses and expressions to a new worksheet. | Modify the sample to also include the evaluated result of each IFERROR formula in the review sheet. | Create a reusable method that returns a DataTable of cell addresses and IFERROR formulas for any given worksheet. | Write a script that scans all worksheets in a workbook for IFERROR usage and consolidates the findings into one summary sheet.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook, scans a worksheet for cells whose formulas contain the IFERROR function, records each cell address and formula, creates a new sheet named "IFERROR Review", writes the collected data with headers, and saves the updated file.
class ExportIfErrorFormulas
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Choose the worksheet to scan (here the first one)
        Worksheet sourceSheet = workbook.Worksheets[0];
        Cells sourceCells = sourceSheet.Cells;

        // Collect addresses and formulas that contain IFERROR
        List<Tuple<string, string>> ifErrorFormulas = new List<Tuple<string, string>>();

        int maxRow = sourceCells.MaxDataRow;
        int maxCol = sourceCells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = sourceCells[row, col];
                if (!string.IsNullOrEmpty(cell.Formula) &&
                    cell.Formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // cell.Name returns the address in A1 style
                    ifErrorFormulas.Add(Tuple.Create(cell.Name, cell.Formula));
                }
            }
        }

        // Add a new worksheet to hold the review list
        Worksheet reviewSheet = workbook.Worksheets.Add("IFERROR Review");
        Cells reviewCells = reviewSheet.Cells;

        // Write header
        reviewCells["A1"].PutValue("Cell Address");
        reviewCells["B1"].PutValue("Formula");

        // Write each collected formula
        for (int i = 0; i < ifErrorFormulas.Count; i++)
        {
            reviewCells[i + 1, 0].PutValue(ifErrorFormulas[i].Item1); // Address
            reviewCells[i + 1, 1].PutValue(ifErrorFormulas[i].Item2); // Formula
        }

        // Save the workbook with the new worksheet
        workbook.Save("output.xlsx");
    }
}
