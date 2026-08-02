// Title: Apply the UNIQUE Dynamic Array Formula with Aspose.Cells for .NET to Auto‑Deduplicate a Column
// Description: This C# example creates a workbook, fills C2:C10 with duplicate and unique strings, inserts the formula =UNIQUE(C2:C10) in D2 using SetDynamicArrayFormula, calculates the sheet, refreshes the spill range, iterates over the distinct values spilled into column D, and saves the file as UniqueDynamicArrayDemo.xlsx.
// Keywords: Aspose.Cells | UNIQUE function | dynamic array formula | C# | .NET | SetDynamicArrayFormula | deduplicate column | spill range | RefreshDynamicArrayFormulas | Excel automation
// Common Searches: Aspose.Cells UNIQUE dynamic array C# example | How to use SetDynamicArrayFormula in Aspose.Cells | Remove duplicate values with UNIQUE in .NET | Refresh dynamic array spill range Aspose.Cells | Get size of UNIQUE spill range programmatically
// Developer Intent: Add a UNIQUE dynamic‑array formula to a worksheet so that duplicate entries in a source column are automatically collapsed into a distinct list that spills into another column.
// Use Cases: Generate a clean list of IDs for reporting without manual filtering. | Create a duplicate‑free source for data‑validation dropdowns. | Prepare a sanitized dataset before exporting to Excel for downstream analysis.
// AI Prompts: Show how to sort the UNIQUE spill results alphabetically after the formula is applied. | Provide code that applies UNIQUE to a numeric range and handles #VALUE! errors. | Explain how to retrieve the number of rows returned by the UNIQUE spill programmatically.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUniqueDynamicArrayDemo
{
    // This C# example creates a workbook, fills C2:C10 with duplicate and unique strings, inserts the formula =UNIQUE(C2:C10) in D2 using SetDynamicArrayFormula, calculates the sheet, refreshes the spill range, iterates over the distinct values spilled into column D, and saves the file as UniqueDynamicArrayDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate column C (C2:C10) with sample data containing duplicates
                // C2:C10 corresponds to row indices 1..9 and column index 2 (zero‑based)
                for (int i = 0; i < 9; i++)
                {
                    // Every third entry is a duplicate value "Dup", others are unique "ValX"
                    string value = (i % 3 == 0) ? "Dup" : $"Val{i}";
                    cells[i + 1, 2].PutValue(value);
                }

                // Set a dynamic array formula in D2 that returns the unique values from C2:C10
                // D2 is row index 1, column index 3 (zero‑based)
                Cell d2 = cells[1, 3];
                d2.SetDynamicArrayFormula("=UNIQUE(C2:C10)", new FormulaParseOptions(), true);

                // Calculate all formulas and refresh dynamic array spill ranges
                workbook.CalculateFormula();
                workbook.RefreshDynamicArrayFormulas(true);

                // Output the spilled results from column D
                // The spill will start at D2 and continue downwards until an empty cell is encountered
                int row = 1;          // start at D2 (row index 1)
                int column = 3;       // column D (index 3)
                Console.WriteLine("Unique values spilled into column D:");
                while (true)
                {
                    Cell current = cells[row, column];
                    // Stop when the cell is blank (no value)
                    if (current == null || current.Value == null || string.IsNullOrEmpty(current.StringValue))
                        break;

                    Console.WriteLine($"{current.Name}: {current.Value}");
                    row++;
                }

                // Ensure the output directory exists
                string outputPath = "UniqueDynamicArrayDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
