using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsDependentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Enable calculation chain to allow dependent tracing
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Ensure all formulas are calculated so the dependency graph is built
            workbook.CalculateFormula();

            // Access the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Specify the cell for which dependents are required (e.g., C1 -> row 0, column 2)
            int targetRow = 0;
            int targetColumn = 2;

            // Retrieve dependents recursively (true = include indirect dependents)
            IEnumerator dependents = cells.GetDependentsInCalculation(targetRow, targetColumn, true);

            Console.WriteLine($"Dependents of cell {cells[targetRow, targetColumn].Name}:");
            if (dependents != null)
            {
                while (dependents.MoveNext())
                {
                    if (dependents.Current is Cell dependentCell)
                    {
                        Console.WriteLine($"- {dependentCell.Name}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No dependents found.");
            }

            // Save the workbook after processing
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}