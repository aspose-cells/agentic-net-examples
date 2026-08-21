// Title: Export Excel to CSV with Formulas as Text Using Aspose.Cells for .NET
// Description: Loads an .xlsx workbook, substitutes each formula cell with its formula string, and saves the workbook as a CSV file so that formulas are retained as plain text instead of evaluated results.
// Keywords: Aspose.Cells CSV export | preserve formulas as text | C# Excel to CSV conversion | formula to string Aspose | save workbook as CSV .NET | extract Excel formulas | convert workbook to CSV | export formulas to CSV
// Common Searches: Aspose.Cells export CSV keep formulas | C# convert Excel to CSV with formulas as text | how to save Excel formulas as text in CSV using Aspose | replace formula with its string before CSV export Aspose.Cells | extract formula strings from Excel with Aspose.Cells
// Developer Intent: Create a CSV file from an Excel workbook where every formula cell is written as its literal formula text.
// Use Cases: Audit or analyze spreadsheet logic by extracting raw formula strings into a CSV file. | Produce CSV reports that display the original calculation expressions for reviewers. | Feed formula text to downstream systems that parse or transform Excel formulas from CSV input.
// AI Prompts: Generate C# code with Aspose.Cells that converts an .xlsx file to .csv and writes each formula cell as its formula string. | Explain how to replace formula values with their textual representation before saving a workbook as CSV using Aspose.Cells for .NET. | Provide a step‑by‑step example that loads a workbook, iterates cells, substitutes formulas with their text, and exports the result to CSV.

using System;
using Aspose.Cells;

namespace WorkbookToCsvWithFormulasAsText
{
    // Loads an .xlsx workbook, substitutes each formula cell with its formula string, and saves the workbook as a CSV file so that formulas are retained as plain text instead of evaluated results.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Path where the CSV output will be saved
            string csvPath = "output.csv";

            // Load the workbook from the source file
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range of the worksheet
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Loop through all cells in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // If the cell contains a formula, replace it with the formula text
                        if (cell.IsFormula)
                        {
                            // Put the formula string as a plain text value
                            cell.PutValue(cell.Formula);
                        }
                    }
                }
            }

            // Save the modified workbook as CSV; formulas are now stored as text strings
            workbook.Save(csvPath, SaveFormat.Csv);

            // Optional: clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook converted to CSV with formulas preserved as text at: {csvPath}");
        }
    }
}
