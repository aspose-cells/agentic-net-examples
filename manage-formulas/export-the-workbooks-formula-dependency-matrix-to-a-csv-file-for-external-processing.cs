// Title: Export Formula Dependency Matrix to CSV with Aspose.Cells for .NET (C#)
// Description: Loads an XLSX workbook, enables the calculation chain, forces formula evaluation, iterates the used range, retrieves recursive dependents via GetDependentsInCalculation, writes SourceCell‑DependentCell pairs to a CSV file, and saves the workbook unchanged.
// Keywords: Aspose.Cells | C# | .NET | Excel formula dependency | GetDependentsInCalculation | CSV export | calculation chain | workbook processing | cell dependents | formula audit
// Common Searches: Aspose.Cells export formula dependencies to CSV | How to get dependent cells in Aspose.Cells .NET | Export Excel calculation chain as CSV file | Retrieve recursive formula dependents using Aspose.Cells | C# code to write Excel cell dependency matrix
// Developer Intent: Generate a CSV file that lists every formula cell and all cells that depend on it.
// Use Cases: Produce an audit report showing which cells influence each formula. | Feed the dependency list into graph‑analysis tools to visualize calculation flow. | Create a lightweight data export for downstream processing while preserving the original workbook.
// AI Prompts: Write C# code with Aspose.Cells that exports a recursive formula dependency matrix to a CSV file, including a header row. | Show how to enable the calculation chain, force formula calculation, and use GetDependentsInCalculation to collect dependent cells. | Explain how to add the worksheet name to each CSV row alongside the cell references.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Loads an XLSX workbook, enables the calculation chain, forces formula evaluation, iterates the used range, retrieves recursive dependents via GetDependentsInCalculation, writes SourceCell‑DependentCell pairs to a CSV file, and saves the workbook unchanged.
class ExportFormulaDependencyMatrix
{
    static void Main()
    {
        // Load the workbook (replace with actual path if needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable calculation chain to allow dependency analysis
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Ensure all formulas are calculated
        workbook.CalculateFormula();

        // Prepare to write CSV output
        using (StreamWriter writer = new StreamWriter("dependency_matrix.csv"))
        {
            // Write CSV header
            writer.WriteLine("SourceCell,DependentCell");

            // Get the first worksheet (adjust if multiple sheets are required)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range to iterate over all cells
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only formula cells (skip empty or value-only cells)
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Get all dependents whose calculated result depends on this cell (recursive)
                        IEnumerator dependents = cells.GetDependentsInCalculation(row, col, true);

                        // If there are dependents, write each pair to the CSV
                        if (dependents != null)
                        {
                            while (dependents.MoveNext())
                            {
                                if (dependents.Current is Cell dependentCell)
                                {
                                    writer.WriteLine($"{cell.Name},{dependentCell.Name}");
                                }
                            }
                        }
                    }
                }
            }
        }

        // Optionally save the workbook after processing (preserves original file)
        workbook.Save("processed.xlsx");
    }
}
