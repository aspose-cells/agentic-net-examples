// Title: Identify and suggest fixes for Excel formulas that reference deleted rows or columns using Aspose.Cells in C#
// AI Prompts: Generate C# code with Aspose.Cells that iterates all worksheets, detects formulas returning #REF! errors, and prints a summary of affected cells. | Update the example to replace each #REF! token in a formula with a placeholder such as "A1" and save the modified workbook. | Create a reusable method that returns a collection of cell addresses and their original formulas where the result is #REF! for further custom processing.
// Common Searches: C# Aspose.Cells find cells with #REF! error after deleting rows | how to programmatically locate broken formula references in an .xlsx file using Aspose.Cells | Aspose.Cells example to list formulas that contain #REF! in a workbook | detect and correct invalid cell references caused by row deletion with Aspose.Cells C#
// Tags: detect #REF! formulas with Aspose.Cells | list cells containing invalid references in .xlsx | correct broken formula references in C# using Aspose.Cells | Aspose.Cells formula validation for deleted rows | scan workbook for #REF! errors in C#

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace FormulaReferenceChecker
{
    // // Uses Aspose.Cells to load a workbook, scans every worksheet, checks each formula cell for a #REF! result indicating a reference to a deleted row or column, and builds suggestion messages describing the problematic cell and formula.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // List to hold correction suggestions
            List<string> suggestions = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the cells collection
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // If the formula result is a #REF! error, the formula references a deleted row/column
                        if (cell.Value != null && cell.Value.ToString() == "#REF!")
                        {
                            string formula = cell.Formula;

                            // Simple detection of #REF! token inside the formula
                            if (formula.Contains("#REF!"))
                            {
                                // Build a suggestion message
                                string suggestion = $"Sheet '{sheet.Name}', Cell {cell.Name}: " +
                                                    $"Formula \"{formula}\" contains a reference to a deleted row/column. " +
                                                    $"Consider removing or updating the #REF! part.";

                                suggestions.Add(suggestion);
                            }
                        }
                    }
                }
            }

            // Output the suggestions
            Console.WriteLine("=== Formula Reference Issues ===");
            if (suggestions.Count == 0)
            {
                Console.WriteLine("No formulas referencing deleted rows/columns were found.");
            }
            else
            {
                foreach (string s in suggestions)
                {
                    Console.WriteLine(s);
                }
            }

            // (Optional) Save the workbook if you made any modifications
            // workbook.Save("output.xlsx");
        }
    }
}
