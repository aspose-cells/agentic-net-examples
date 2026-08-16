// Title: Aspose.Cells C# – Export Formula Dependent Cells to CSV with GetDependents
// Description: Creates a workbook, adds formulas that reference cell A1, calculates all formulas, retrieves every dependent cell (direct, indirect, and cross‑worksheet) using the GetDependents method, and writes the cell addresses to a CSV file with a header row.
// Keywords: Aspose.Cells | GetDependents | C# | CSV export | dependent cells | formula audit | cross‑worksheet dependencies | impact analysis | cell address list
// Common Searches: Aspose.Cells get dependent cells C# | export formula dependents to CSV | GetDependents cross worksheet example | list cells affected by a source cell Aspose.Cells | how to save dependent cell names as CSV
// Developer Intent: Extract all cells that rely on a specific source cell and save their addresses in a CSV file for reporting or analysis.
// Use Cases: Produce an audit report showing which cells will recalculate when a source cell changes. | Create a CSV inventory of dependent cells for downstream data pipelines or documentation. | Perform impact analysis across multiple worksheets before modifying a critical formula.
// AI Prompts: Generate C# code that uses Aspose.Cells to find dependents of cell B2 across all worksheets and writes the results to a JSON file. | Provide a reusable method that accepts a Workbook and a cell reference, returns a list of dependent cell names, and exports them to CSV with error handling and logging.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDependentsExport
{
    // Creates a workbook, adds formulas that reference cell A1, calculates all formulas, retrieves every dependent cell (direct, indirect, and cross‑worksheet) using the GetDependents method, and writes the cell addresses to a CSV file with a header row.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data: set a source value and formulas that depend on it
            cells["A1"].PutValue(10);               // source cell
            cells["B1"].Formula = "=A1*2";          // direct dependent
            cells["C1"].Formula = "=A1+B1";         // indirect dependent
            cells["D1"].Formula = "=B1*3";          // direct dependent
            cells["F4"].Formula = "=A1*5";          // direct dependent on another sheet (if any)

            // Calculate all formulas so that dependents are recognized
            workbook.CalculateFormula();

            // Retrieve all dependents of cell A1 (row 0, column 0), including indirect ones
            // isAll = false -> only current worksheet; set true to include other worksheets
            Cell[] dependents = cells.GetDependents(true, 0, 0);

            // Prepare CSV file path
            string csvPath = "A1_Dependents.csv";

            // Export dependents to CSV (one cell name per line)
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write header
                writer.WriteLine("DependentCellName");

                // Write each dependent cell name
                foreach (Cell dep in dependents)
                {
                    writer.WriteLine(dep.Name);
                }
            }

            // Save the workbook (optional, just to keep the sample workbook)
            workbook.Save("DependentsSample.xlsx");

            Console.WriteLine($"Dependents of A1 have been exported to '{csvPath}'.");
        }
    }
}
