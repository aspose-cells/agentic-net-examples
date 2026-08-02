// Title: C# – Extract formulas from a named range to a text file with Aspose.Cells
// Description: Loads an Excel workbook, finds a defined name (named range), iterates through every cell in all areas of the range, captures the formula text of each formula cell, and writes the cell address and formula to a plain‑text file. Includes checks for missing files or undefined names and supports multi‑area ranges.
// Keywords: Aspose.Cells | C# extract formulas | named range | export formulas to txt | Excel formula extraction .NET | multi‑area named range | cell.Formula | Workbook.Load | write formulas file
// Common Searches: how to get formula strings from a named range using Aspose.Cells | export Excel formulas to a .txt file in C# | iterate over multi‑area named ranges and write formulas Aspose.Cells | save cell formulas to a text file .NET | Aspose.Cells read formula text from named range
// Developer Intent: Read the formula text of every cell inside a specific named range and save each cell address with its formula to a plain‑text file.
// Use Cases: Create an audit report that lists all formulas used in a defined range for compliance checks. | Export formulas from a source workbook to compare with a target system during migration. | Debug complex worksheets by dumping formulas from a named range into a readable file.
// AI Prompts: Generate C# code with Aspose.Cells that extracts formulas from a named range and writes them to a CSV file. | Show how to log the address and formula of each cell in a multi‑area named range using Aspose.Cells. | Provide a robust example that handles a missing named range and outputs a clear error message.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, finds a defined name (named range), iterates through every cell in all areas of the range, captures the formula text of each formula cell, and writes the cell address and formula to a plain‑text file. Includes checks for missing files or undefined names and supports multi‑area ranges.
class ExtractFormulasFromNamedRange
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Path to the output text file
        string txtPath = "formulas.txt";

        // Name of the defined range to process
        string rangeName = "MyRange";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file '{excelPath}' not found.");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(excelPath);

            // Retrieve the Name object that represents the named range
            Name namedRange = workbook.Worksheets.Names[rangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range '{rangeName}' not found.");
                return;
            }

            // Get all ranges referred by the name (handles multi‑area names)
            Aspose.Cells.Range[] ranges = namedRange.GetRanges();

            using (StreamWriter writer = new StreamWriter(txtPath))
            {
                // Iterate through each range
                foreach (Aspose.Cells.Range range in ranges)
                {
                    // Iterate through each cell inside the current range
                    for (int r = range.FirstRow; r < range.FirstRow + range.RowCount; r++)
                    {
                        for (int c = range.FirstColumn; c < range.FirstColumn + range.ColumnCount; c++)
                        {
                            Cell cell = workbook.Worksheets[range.Worksheet.Index].Cells[r, c];
                            if (cell.IsFormula)
                            {
                                // Write cell address and its formula text to the file
                                writer.WriteLine($"{cell.Name}: {cell.Formula}");
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"Formulas from named range '{rangeName}' have been written to '{txtPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
