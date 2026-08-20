// Title: C# – Export IFERROR and IFNA Formulas from an Excel Workbook with Aspose.Cells
// Description: Loads a workbook, scans every used cell for IFERROR or IFNA functions (case‑insensitive), records the sheet name, cell address, and formula, and writes the results to a new Excel file.
// Keywords: Aspose.Cells IFERROR export | Aspose.Cells IFNA extraction | C# find error‑handling formulas | export formulas to new workbook .NET | audit Excel formulas Aspose | list IFERROR cells C# | extract IFNA formulas Aspose.Cells
// Common Searches: how to locate IFERROR formulas using Aspose.Cells | export cells with IFNA to another workbook C# | scan workbook for error handling functions Aspose | save extracted formulas with sheet and address | Aspose.Cells list formulas containing IFERROR
// Developer Intent: Detect every IFERROR or IFNA formula in a workbook and export its sheet, cell reference, and expression to a separate Excel file.
// Use Cases: Generate an audit report of all error‑handling formulas for quality assurance. | Create a lightweight workbook that contains only the identified formulas for debugging. | Log formula details to the console or a file for documentation and further analysis.
// AI Prompts: Write C# code with Aspose.Cells that extracts IFERROR/IFNA formulas and saves them as CSV. | Show how to extend the sample to also capture ISERROR or ISNA functions. | Suggest performance‑optimisation techniques for scanning very large workbooks for specific functions using Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaExport
{
    // Loads a workbook, scans every used cell for IFERROR or IFNA functions (case‑insensitive), records the sheet name, cell address, and formula, and writes the results to a new Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook that contains formulas
            string sourcePath = "SourceWorkbook.xlsx";

            // Load the source workbook (load rule)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // List to hold information about formulas that use IFERROR or IFNA
            List<(string SheetName, string CellName, string Formula)> errorHandlingFormulas = new List<(string, string, string)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Only consider cells that actually have a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Check for IFERROR or IFNA (case‑insensitive)
                        string formulaUpper = cell.Formula.ToUpperInvariant();
                        if (formulaUpper.Contains("IFERROR") || formulaUpper.Contains("IFNA"))
                        {
                            // Store sheet name, cell address and the formula text
                            errorHandlingFormulas.Add((sheet.Name, cell.Name, cell.Formula));
                        }
                    }
                }
            }

            // Create a new workbook to export the collected formulas (create rule)
            Workbook exportWorkbook = new Workbook();
            Worksheet exportSheet = exportWorkbook.Worksheets[0];
            Cells exportCells = exportSheet.Cells;

            // Write header row
            exportCells["A1"].PutValue("Sheet");
            exportCells["B1"].PutValue("Cell");
            exportCells["C1"].PutValue("Formula");

            // Populate the export sheet with the collected data
            int rowIndex = 1; // zero‑based index; start after header
            foreach (var item in errorHandlingFormulas)
            {
                exportCells[rowIndex, 0].PutValue(item.SheetName);
                exportCells[rowIndex, 1].PutValue(item.CellName);
                exportCells[rowIndex, 2].PutValue(item.Formula);
                rowIndex++;
            }

            // Save the export workbook (save rule)
            string exportPath = "ExportedErrorHandlingFormulas.xlsx";
            exportWorkbook.Save(exportPath);

            // Optional: also output to console for quick verification
            Console.WriteLine($"Found {errorHandlingFormulas.Count} formulas with IFERROR/IFNA.");
            foreach (var item in errorHandlingFormulas)
            {
                Console.WriteLine($"{item.SheetName}!{item.CellName}: {item.Formula}");
            }
        }
    }
}
