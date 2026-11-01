// Title: Find Excel cells whose formulas reference a specific named range using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells find cells with formulas that reference a named range | how to locate formulas that use a particular named range in an .xlsx file with Aspose.Cells | retrieve cell addresses of formulas containing MyNamedRange using Aspose.Cells API | scan worksheet for named range references in formulas with Aspose.Cells .NET | list all cells referencing a specific named range in Excel using C# Aspose.Cells
// Tags: Aspose.Cells find formula references | C# search formulas by named range | Excel .xlsx named range usage detection | iterate worksheet cells Aspose.Cells | collect cell addresses with specific formula text

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace FindFormulaReferencesApp
{
    // The example loads an .xlsx workbook with Aspose.Cells, iterates over all used cells in the first worksheet, checks each cell's formula for the target named range (case‑insensitive), gathers matching cells, prints their addresses and formulas, and optionally saves the workbook.
    class FindFormulaReferences
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];
                const string namedRange = "MyNamedRange";

                List<Cell> cellsWithReference = new List<Cell>();

                // Scan used cells for formulas containing the named range
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxColumn = worksheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];
                        if (cell.IsFormula && cell.Formula != null &&
                            cell.Formula.IndexOf(namedRange, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            cellsWithReference.Add(cell);
                        }
                    }
                }

                // Output addresses and formulas of found cells
                foreach (Cell cell in cellsWithReference)
                {
                    Console.WriteLine($"Cell {cell.Name} contains formula: {cell.Formula}");
                }

                // Save the workbook (optional)
                const string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
