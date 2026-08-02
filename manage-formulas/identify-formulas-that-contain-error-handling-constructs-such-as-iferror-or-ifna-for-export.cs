// Title: Export IFERROR and IFNA Formulas from Excel with Aspose.Cells for .NET
// Description: Loads a workbook, scans every worksheet and cell for formulas that contain IFERROR or IFNA (case‑insensitive), records each cell address and its formula, writes the results to a new workbook with headers, and saves the file.
// Keywords: Aspose.Cells C# | export IFERROR formulas | extract IFNA formulas | list error‑handling formulas | iterate worksheets Aspose | search formulas Excel | Excel formula extraction | Aspose.Cells example GitHub | C# Excel automation | error handling formula export
// Common Searches: How to extract IFERROR formulas using Aspose.Cells | Export cells with IFNA to a new workbook C# | List all error handling formulas in an Excel file Aspose | Aspose.Cells find formulas containing IFERROR or IFNA | C# code to export formulas that handle errors
// Developer Intent: Find and export every cell that uses IFERROR or IFNA in a source workbook into a separate Excel file.
// Use Cases: Create an audit report of all error‑handling formulas for compliance checks. | Provide developers or QA teams with a lightweight workbook that contains only the problematic formulas. | Pre‑process a workbook by extracting IFERROR/IFNA formulas before running bulk data transformations.
// AI Prompts: Generate C# code with Aspose.Cells that collects cells whose formulas contain IFERROR or IFNA and writes their addresses and formulas to a new workbook. | Extend the sample to also capture formulas that use ISERROR or IFERROR with nested functions. | Explain how to make the formula search case‑insensitive and include merged cells in the export.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaExport
{
    // Loads a workbook, scans every worksheet and cell for formulas that contain IFERROR or IFNA (case‑insensitive), records each cell address and its formula, writes the results to a new workbook with headers, and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "ErrorHandlingFormulas.xlsx";

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // List to hold cells that contain IFERROR or IFNA
            List<(string Address, string Formula)> errorHandlingFormulas = new List<(string, string)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all used cells
                foreach (Cell cell in cells)
                {
                    // Check if the cell has a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Look for IFERROR or IFNA (case‑insensitive)
                        if (cell.Formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            cell.Formula.IndexOf("IFNA", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Store the cell address (e.g., A1) and its formula
                            errorHandlingFormulas.Add((cell.Name, cell.Formula));
                        }
                    }
                }
            }

            // Create a new workbook to export the results
            Workbook exportWorkbook = new Workbook();
            Worksheet exportSheet = exportWorkbook.Worksheets[0];
            Cells exportCells = exportSheet.Cells;

            // Write header
            exportCells["A1"].PutValue("Cell Address");
            exportCells["B1"].PutValue("Formula");

            // Populate the export sheet with collected formulas
            int rowIndex = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (var item in errorHandlingFormulas)
            {
                exportCells[rowIndex, 0].PutValue(item.Address);
                exportCells[rowIndex, 1].PutValue(item.Formula);
                rowIndex++;
            }

            // Save the export workbook
            exportWorkbook.Save(outputPath);
        }
    }
}
