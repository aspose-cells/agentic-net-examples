using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDependentExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string excelPath = "InputWorkbook.xlsx";

            // Path to the output CSV file
            string csvPath = "Dependents.csv";

            // Address of the cell whose dependents we want to export (e.g., "A1")
            string targetCellAddress = "A1";

            // Load the workbook (creation rule)
            Workbook workbook = new Workbook(excelPath);

            // Ensure all formulas are calculated before tracing dependents
            workbook.CalculateFormula();

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Get the target cell
            Cell targetCell = cells[targetCellAddress];

            // Retrieve all dependents (including indirect) of the target cell
            // true => check other worksheets as well
            Cell[] dependents = targetCell.GetDependents(true);

            // Write dependents information to CSV
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Header row
                writer.WriteLine("DependentCell,Formula");

                // Iterate through each dependent cell
                foreach (Cell dep in dependents)
                {
                    // If the dependent cell contains a formula, use it; otherwise leave empty
                    string formula = dep.IsFormula ? dep.Formula : string.Empty;

                    // Escape commas in formula if any
                    if (formula.Contains(","))
                    {
                        formula = $"\"{formula}\"";
                    }

                    writer.WriteLine($"{dep.Name},{formula}");
                }
            }

            // Save the workbook if any modifications were made (save rule)
            workbook.Save("InputWorkbook_Saved.xlsx");

            Console.WriteLine($"Exported {dependents.Length} dependents of {targetCellAddress} to '{csvPath}'.");
        }
    }
}