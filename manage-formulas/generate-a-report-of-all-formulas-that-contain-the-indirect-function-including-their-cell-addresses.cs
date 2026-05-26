using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsIndirectReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(sourcePath);

            // List to hold addresses of cells containing INDIRECT in their formulas
            List<string> indirectCells = new List<string>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Look for the INDIRECT function (case‑insensitive)
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Record the full address: SheetName!CellName
                            indirectCells.Add($"{sheet.Name}!{cell.Name}");
                        }
                    }
                }
            }

            // Create a new worksheet to hold the report
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "IndirectReport";

            // Write header
            reportSheet.Cells[0, 0].PutValue("Cell Address");
            reportSheet.Cells[0, 1].PutValue("Formula");

            // Populate the report
            for (int i = 0; i < indirectCells.Count; i++)
            {
                // Split the stored address back into sheet and cell name
                string[] parts = indirectCells[i].Split('!');
                string sheetName = parts[0];
                string cellName = parts[1];

                // Retrieve the original formula
                Cell originalCell = workbook.Worksheets[sheetName].Cells[cellName];
                string formula = originalCell.Formula;

                // Write to the report sheet
                reportSheet.Cells[i + 1, 0].PutValue(indirectCells[i]); // full address
                reportSheet.Cells[i + 1, 1].PutValue(formula);
            }

            // Save the workbook with the new report sheet (lifecycle rule: save)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            // Optional: also write the report to console
            Console.WriteLine("Cells containing INDIRECT function:");
            foreach (string address in indirectCells)
            {
                Console.WriteLine(address);
            }
        }
    }
}