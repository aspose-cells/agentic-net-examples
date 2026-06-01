using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

class ExportFormulaDependencyMatrix
{
    static void Main()
    {
        try
        {
            // Path to the source Excel workbook
            string workbookPath = "InputWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
                return;
            }

            // Path to the output CSV file
            string csvPath = "FormulaDependencyMatrix.csv";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Enable calculation chain to allow dependency queries
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Ensure all formulas are calculated so the dependency chain is built
            workbook.CalculateFormula();

            // Prepare a StreamWriter for the CSV output
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("SourceCell,DependentCell");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all used cells in the worksheet
                    foreach (Cell cell in cells)
                    {
                        // Process only cells that contain a formula
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            // Get all dependents (recursive) for the current cell
                            // Correct signature: GetDependentsInCalculation(int row, int column, bool recursive)
                            IEnumerator dependents = cells.GetDependentsInCalculation(cell.Row, cell.Column, true);

                            if (dependents != null)
                            {
                                while (dependents.MoveNext())
                                {
                                    if (dependents.Current is Cell dependentCell)
                                    {
                                        // Write a CSV line: source cell name, dependent cell name (including sheet name)
                                        string source = $"{sheet.Name}!{cell.Name}";
                                        string dependent = $"{dependentCell.Worksheet.Name}!{dependentCell.Name}";
                                        writer.WriteLine($"{source},{dependent}");
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"Formula dependency matrix exported to '{csvPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}