using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class TracingDependentsDemo
    {
        public static void Run()
        {
            // Load an existing workbook (replace with actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Enable calculation chain to allow dependent tracing
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Perform full calculation so the chain is built
            workbook.CalculateFormula();

            // Get the worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example: trace dependents of cell C1 (row 0, column 2) recursively
            IEnumerator dependents = cells.GetDependentsInCalculation(0, 2, true);

            Console.WriteLine("Dependents of cell C1:");
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

            // Save the workbook (optional, can be omitted if only tracing is needed)
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TracingDependentsDemo.Run();
        }
    }
}