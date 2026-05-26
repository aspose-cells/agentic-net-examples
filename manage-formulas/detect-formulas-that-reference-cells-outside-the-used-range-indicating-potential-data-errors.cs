using System;
using System.Collections.Generic;
using Aspose.Cells;

class DetectOutOfRangeFormulas
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Work with the first worksheet (adjust as needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;      // zero‑based index of the last used row
        int maxCol = cells.MaxDataColumn;   // zero‑based index of the last used column

        List<string> problematic = new List<string>();

        // Iterate through all cells in the worksheet
        foreach (Cell cell in cells)
        {
            // Process only formula cells
            if (!cell.IsFormula) continue;

            // Get all precedent references from the formula
            ReferredAreaCollection precedents = cell.GetPrecedents();
            if (precedents == null) continue;

            foreach (ReferredArea area in precedents)
            {
                // Skip external links – they are not part of the current worksheet's used range
                if (area.IsExternalLink) continue;

                bool outOfRange = false;

                // Handle whole‑column or whole‑row references
                if (area.IsEntireColumn)
                {
                    // Column index beyond the last used column
                    if (area.StartColumn > maxCol) outOfRange = true;
                }
                else if (area.IsEntireRow)
                {
                    // Row index beyond the last used row
                    if (area.StartRow > maxRow) outOfRange = true;
                }
                else
                {
                    // Normal range – check start and end positions
                    if (area.StartRow > maxRow || area.EndRow > maxRow ||
                        area.StartColumn > maxCol || area.EndColumn > maxCol)
                    {
                        outOfRange = true;
                    }
                }

                if (outOfRange)
                {
                    // Record the formula cell and the offending reference
                    problematic.Add($"{cell.Name} (Formula: {cell.Formula}) -> {area}");
                    // One offending reference is enough to flag the cell
                    break;
                }
            }
        }

        // Output the results
        Console.WriteLine("Formulas referencing cells outside the used range:");
        foreach (string entry in problematic)
        {
            Console.WriteLine(entry);
        }

        // Optionally save the workbook (e.g., after marking cells)
        workbook.Save("output.xlsx");
    }
}