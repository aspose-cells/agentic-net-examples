// Title: Export Excel Formulas to CSV with Cell Addresses Using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook with Aspose.Cells, scans every worksheet, extracts cells that contain formulas, and writes a CSV file that records the worksheet name, the A1‑style cell address, and the formula (quotes escaped).
// Keywords: Aspose.Cells export formulas CSV | C# extract Excel formulas | list cell formulas Aspose.Cells | write formulas to CSV .NET | iterate worksheets Aspose.Cells | Excel formula extraction C# | CSV export Aspose.Cells
// Common Searches: how to export all formulas from Excel to CSV using Aspose.Cells | c# code to get cell address and formula Aspose.Cells | save Excel formulas with worksheet name to CSV | extract formulas from workbook with Aspose.Cells .NET | csv export of Excel formulas c#
// Developer Intent: Generate a CSV file that lists every formula in an Excel workbook together with its worksheet and cell address.
// Use Cases: Create an audit report of all calculations for compliance checks. | Document spreadsheet logic for knowledge‑transfer or migration projects. | Feed formula data into a tool that converts Excel models to another platform.
// AI Prompts: Write C# code with Aspose.Cells to export all workbook formulas to a CSV file, handling quote escaping. | Adapt the sample to export formulas only from a user‑specified worksheet. | Add a column that flags whether each formula is an array formula. | Include the formula result value alongside the formula text in the CSV output.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaExport
{
    // Loads an Excel workbook with Aspose.Cells, scans every worksheet, extracts cells that contain formulas, and writes a CSV file that records the worksheet name, the A1‑style cell address, and the formula (quotes escaped).
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook
            string workbookPath = "input.xlsx";

            // Path for the output CSV file
            string csvPath = "formulas.csv";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Prepare a StreamWriter for the CSV output
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("Worksheet,CellAddress,Formula");

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the Cells collection for the current worksheet
                    Cells cells = sheet.Cells;

                    // Iterate through all used cells in the worksheet
                    foreach (Cell cell in cells)
                    {
                        // Check if the cell contains a formula
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            // Get the cell address in A1 style
                            string address = cell.Name;

                            // Write a CSV line: Worksheet name, cell address, formula text
                            // Escape double quotes in the formula by doubling them
                            string escapedFormula = cell.Formula.Replace("\"", "\"\"");
                            writer.WriteLine($"{sheet.Name},{address},\"{escapedFormula}\"");
                        }
                    }
                }
            }

            Console.WriteLine($"Formulas have been exported to '{csvPath}'.");
        }
    }
}
