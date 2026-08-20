// Title: Find Excel formulas that reference cells outside the used range using Aspose.Cells for .NET
// Description: Loads an Excel workbook, determines each worksheet's used range with MaxDataRow/MaxDataColumn, scans all formula cells, examines their precedent areas via GetPrecedents, and flags any reference that lies beyond the used rows or columns (including whole‑row, whole‑column, single‑cell, and range references). The program outputs the offending cell addresses and formulas and can save the workbook.
// Keywords: Aspose.Cells | .NET | detect out of range formula references | Excel formula validation | used range detection | MaxDataRow | MaxDataColumn | GetPrecedents | precedent area analysis | invalid cell reference | data quality audit | automated Excel checks
// Common Searches: Aspose.Cells find formulas referencing cells outside used range | detect out‑of‑range precedent areas in Excel with .NET | list formula cells that point to non‑existent rows or columns | validate Excel formulas using Aspose.Cells GetPrecedents | how to flag formulas that reference empty rows in Aspose.Cells
// Developer Intent: Identify and list all formula cells that reference rows or columns beyond the worksheet's used range.
// Use Cases: Generate a quality‑control report of potentially erroneous formulas before workbook distribution. | Automate data‑integrity checks in ETL pipelines by flagging formulas that point to empty rows or columns. | Integrate out‑of‑range formula detection into CI/CD builds to prevent publishing faulty spreadsheets.
// AI Prompts: Create a method that returns a collection of Cell objects whose formulas reference rows or columns beyond MaxDataRow/MaxDataColumn. | Enhance the sample to apply a red background style to each offending cell after detection. | Write a unit test suite that verifies out‑of‑range detection for single‑cell, whole‑row, and whole‑column references.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaReferenceCheck
{
    // Loads an Excel workbook, determines each worksheet's used range with MaxDataRow/MaxDataColumn, scans all formula cells, examines their precedent areas via GetPrecedents, and flags any reference that lies beyond the used rows or columns (including whole‑row, whole‑column, single‑cell, and range references). The program outputs the offending cell addresses and formulas and can save the workbook.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // List to hold cells with out‑of‑range references
            List<Cell> cellsWithInvalidRefs = new List<Cell>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range of the worksheet
                int maxRow = cells.MaxDataRow;      // zero‑based index of last used row
                int maxCol = cells.MaxDataColumn;   // zero‑based index of last used column

                // Iterate over all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        // Get all precedent areas referenced by the formula
                        ReferredAreaCollection precedents = cell.GetPrecedents();
                        if (precedents == null) continue;

                        foreach (ReferredArea area in precedents)
                        {
                            // Skip external links – they are not part of the current worksheet's used range
                            if (area.IsExternalLink) continue;

                            // Check if the referenced area lies outside the used range
                            bool outOfRange = false;

                            // Single cell reference
                            if (!area.IsArea)
                            {
                                if (area.StartRow > maxRow || area.StartColumn > maxCol)
                                    outOfRange = true;
                            }
                            else // Range reference
                            {
                                // Entire column reference
                                if (area.IsEntireColumn)
                                {
                                    if (area.StartColumn > maxCol)
                                        outOfRange = true;
                                }
                                // Entire row reference
                                else if (area.IsEntireRow)
                                {
                                    if (area.StartRow > maxRow)
                                        outOfRange = true;
                                }
                                else
                                {
                                    // Normal range
                                    if (area.EndRow > maxRow || area.EndColumn > maxCol)
                                        outOfRange = true;
                                }
                            }

                            if (outOfRange)
                            {
                                cellsWithInvalidRefs.Add(cell);
                                // No need to check other areas for this cell
                                break;
                            }
                        }
                    }
                }
            }

            // Output results
            Console.WriteLine("Cells with formulas referencing outside the used range:");
            foreach (Cell c in cellsWithInvalidRefs)
            {
                Console.WriteLine($"{c.Name} (Sheet: {c.Worksheet.Name}) -> Formula: {c.Formula}");
            }

            // Optionally, save the workbook (if any modifications were made)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
