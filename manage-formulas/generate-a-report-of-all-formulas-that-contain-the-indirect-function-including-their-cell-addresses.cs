// Title: C# – Aspose.Cells: Generate a worksheet report of all INDIRECT formulas with cell addresses
// Description: Loads an Excel file, iterates through every worksheet and used cell, detects formulas that contain the INDIRECT function (case‑insensitive), records each cell's address and formula, writes the results to a new sheet named "IndirectReport", and saves the workbook. Ideal for auditing dynamic references in Excel workbooks using Aspose.Cells for .NET.
// Keywords: Aspose.Cells INDIRECT report C# | list INDIRECT formulas Aspose.Cells | extract formula addresses .NET | Excel audit dynamic references | C# generate formula report
// Common Searches: Aspose.Cells list all INDIRECT formulas in a workbook | C# code to find formulas containing INDIRECT | Create a report of Excel formulas with Aspose.Cells | How to extract cell addresses of INDIRECT functions using .NET | Generate a summary sheet of dynamic references in Excel
// Developer Intent: Find every formula that uses INDIRECT, capture its address, and output the data to a new worksheet.
// Use Cases: Debugging: locate all dynamic references created with INDIRECT across a workbook. | Compliance: document every INDIRECT formula for audit or review purposes. | Performance analysis: consolidate indirect references to assess calculation impact.
// AI Prompts: Write C# code with Aspose.Cells that searches for a specified function in formulas and writes the cell address and formula to a new sheet. | Suggest a more efficient way to enumerate only used cells when generating the INDIRECT formula report. | Explain how to extend the report to include the worksheet name alongside each cell address.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsIndirectReport
{
    // Loads an Excel file, iterates through every worksheet and used cell, detects formulas that contain the INDIRECT function (case‑insensitive), records each cell's address and formula, writes the results to a new sheet named "IndirectReport", and saves the workbook. Ideal for auditing dynamic references in Excel workbooks using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Path to the output workbook (will contain the report)
            string outputPath = "output_with_indirect_report.xlsx";

            // Load the workbook (using the provided load rule)
            Workbook workbook = new Workbook(inputPath);

            // Prepare a list to hold report entries (cell address and formula)
            List<(string Address, string Formula)> indirectFormulas = new List<(string, string)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the Cells collection of the current worksheet
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        // Get the formula text
                        string formula = cell.Formula;

                        // Look for the INDIRECT function (case‑insensitive)
                        if (!string.IsNullOrEmpty(formula) && 
                            formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Store the cell address (e.g., A1) and its formula
                            indirectFormulas.Add((cell.Name, formula));
                        }
                    }
                }
            }

            // Create a new worksheet to hold the report (using the create rule)
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "IndirectReport";

            // Write header
            reportSheet.Cells[0, 0].PutValue("Cell Address");
            reportSheet.Cells[0, 1].PutValue("Formula");

            // Populate the report
            for (int i = 0; i < indirectFormulas.Count; i++)
            {
                reportSheet.Cells[i + 1, 0].PutValue(indirectFormulas[i].Address);
                reportSheet.Cells[i + 1, 1].PutValue(indirectFormulas[i].Formula);
            }

            // Save the workbook (using the provided save rule)
            workbook.Save(outputPath);

            // Optional: display result in console
            Console.WriteLine($"Found {indirectFormulas.Count} formulas containing INDIRECT.");
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
