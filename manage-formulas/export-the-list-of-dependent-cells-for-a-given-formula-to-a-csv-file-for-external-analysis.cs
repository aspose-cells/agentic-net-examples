// Title: Export Formula Dependent Cells to CSV Using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, calculates all formulas, identifies every cell that directly or indirectly depends on a specified address with GetDependents, and writes each dependent's address and formula to a CSV file with proper escaping. An optional save of the workbook is also demonstrated.
// Keywords: Aspose.Cells | C# | .NET | GetDependents | dependent cells | Excel formula tracing | export to CSV | cell dependency report | Excel automation | Aspose.Cells example
// Common Searches: Aspose.Cells export dependent cells to CSV | C# list cells that depend on A1 using Aspose.Cells | Get indirect dependents of a formula in .NET | Trace formula dependencies and save as CSV | How to write dependent cell addresses to a file with Aspose.Cells
// Developer Intent: Retrieve all cells that rely on a given formula cell and generate a CSV file containing their addresses and formulas.
// Use Cases: Produce a dependency matrix for a budgeting model by exporting all cells linked to a key input. | Assess the impact of changing a parameter by analyzing its downstream formulas in an external analytics tool. | Feed dependent‑cell information into a custom validation pipeline that operates outside of Excel.
// AI Prompts: Write C# code that uses Aspose.Cells to collect all dependent cells of a target address and export their names and formulas to a CSV file. | Show how to filter the dependent list so only cells containing formulas are written to the CSV. | Explain how to extend the export to include separate columns for worksheet name, cell address, and formula.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDependentExport
{
    // Loads an Excel workbook, calculates all formulas, identifies every cell that directly or indirectly depends on a specified address with GetDependents, and writes each dependent's address and formula to a CSV file with proper escaping. An optional save of the workbook is also demonstrated.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file, output CSV file and the cell to analyze (e.g., "A1")
            string excelPath = "InputWorkbook.xlsx";
            string csvPath = "Dependents.csv";
            string targetCellAddress = "A1";

            // ---------- Create / Load ----------
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(excelPath);
            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Ensure all formulas are calculated before tracing dependents
            workbook.CalculateFormula();

            // Convert the target cell address to zero‑based row and column indices
            int targetRow, targetColumn;
            CellsHelper.CellNameToIndex(targetCellAddress, out targetRow, out targetColumn);

            // ---------- Get Dependents ----------
            // Retrieve all cells that depend on the target cell (including indirect dependents)
            Cell[] dependents = cells.GetDependents(true, targetRow, targetColumn);

            // ---------- Export to CSV ----------
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write header
                writer.WriteLine("DependentCellName,Formula");

                // Write each dependent cell's name and its formula (if any)
                foreach (Cell dep in dependents)
                {
                    string formula = dep.IsFormula ? dep.Formula : string.Empty;
                    writer.WriteLine($"{dep.Name},{EscapeCsv(formula)}");
                }
            }

            // ---------- Save ----------
            // (Optional) Save the workbook if any modifications were made
            workbook.Save("ModifiedWorkbook.xlsx");
        }

        // Helper to escape commas and quotes in CSV fields
        private static string EscapeCsv(string field)
        {
            if (field == null) return string.Empty;
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            {
                field = field.Replace("\"", "\"\"");
                return $"\"{field}\"";
            }
            return field;
        }
    }
}
