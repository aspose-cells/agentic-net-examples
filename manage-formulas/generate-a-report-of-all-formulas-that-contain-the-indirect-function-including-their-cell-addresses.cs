// Title: C# – List all INDIRECT formulas with cell addresses using Aspose.Cells
// Description: Loads an Excel file, scans every worksheet for formulas that contain the INDIRECT function (case‑insensitive), captures the sheet name, cell address and formula, writes the data to a new worksheet called IndirectReport, auto‑fits columns and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | INDIRECT | Excel formula audit | list formulas | cell address | report generation | search formulas | enumerate cells | dynamic references
// Common Searches: list INDIRECT formulas Aspose.Cells C# | how to find cells with INDIRECT function in Excel using .NET | generate formula audit workbook with Aspose.Cells | C# code to extract formulas containing INDIRECT | Aspose.Cells enumerate formulas by keyword
// Developer Intent: Create a workbook that enumerates every cell using the INDIRECT function, showing its sheet, address and formula.
// Use Cases: Audit dynamic references before migrating a spreadsheet to another platform. | Identify performance‑critical INDIRECT formulas for optimization. | Provide a quick reference for developers to locate and review indirect logic.
// AI Prompts: Write C# code with Aspose.Cells that extracts all formulas containing INDIRECT and writes them to a new worksheet. | Give a method that filters cells by a keyword in their formula and returns sheet name, address, and formula. | Explain how to extend the report to include the evaluated value of each INDIRECT formula.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsIndirectReport
{
    // Loads an Excel file, scans every worksheet for formulas that contain the INDIRECT function (case‑insensitive), captures the sheet name, cell address and formula, writes the data to a new worksheet called IndirectReport, auto‑fits columns and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with actual file path)
            string sourcePath = "source.xlsx";

            // Path to the output workbook containing the report
            string outputPath = "IndirectFormulasReport.xlsx";

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // List to hold report entries
            List<(string SheetName, string CellName, string Formula)> indirectFormulas = new List<(string, string, string)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain data or formulas
                foreach (Cell cell in cells)
                {
                    // Check if the cell has a formula and contains the INDIRECT function (case‑insensitive)
                    if (!string.IsNullOrEmpty(cell.Formula) &&
                        cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        indirectFormulas.Add((sheet.Name, cell.Name, cell.Formula));
                    }
                }
            }

            // Add a new worksheet for the report
            int reportIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportIndex];
            reportSheet.Name = "IndirectReport";

            // Write header row
            reportSheet.Cells["A1"].PutValue("Worksheet");
            reportSheet.Cells["B1"].PutValue("Cell");
            reportSheet.Cells["C1"].PutValue("Formula");

            // Populate report data
            int row = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (var entry in indirectFormulas)
            {
                reportSheet.Cells[row, 0].PutValue(entry.SheetName);
                reportSheet.Cells[row, 1].PutValue(entry.CellName);
                reportSheet.Cells[row, 2].PutValue(entry.Formula);
                row++;
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the workbook with the report
            workbook.Save(outputPath);
        }
    }
}
