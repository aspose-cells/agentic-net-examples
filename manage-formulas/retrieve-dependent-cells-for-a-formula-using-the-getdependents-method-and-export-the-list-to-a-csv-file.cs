using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDependentsToCsv
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data and formulas
            cells["A1"].PutValue(10);
            cells["B1"].Formula = "=A1*2";
            cells["C1"].Formula = "=A1+B1";
            cells["D1"].Formula = "=B1*3";
            cells["F4"].Formula = "=A1*5";

            // Calculate all formulas so that dependents are recognized
            workbook.CalculateFormula();

            // Get direct dependents of cell A1 (row 0, column 0)
            Cell[] directDependents = cells.GetDependents(false, 0, 0);

            // Get all (including indirect) dependents of cell A1
            Cell[] allDependents = cells.GetDependents(true, 0, 0);

            // Export direct dependents to CSV
            ExportDependentsToCsv("DirectDependents.csv", directDependents);

            // Export all dependents to CSV
            ExportDependentsToCsv("AllDependents.csv", allDependents);

            // Save the workbook (optional, just to demonstrate lifecycle rule)
            workbook.Save("DependentsDemo.xlsx");
        }

        // Helper method to write an array of Cell objects to a CSV file
        private static void ExportDependentsToCsv(string filePath, Cell[] dependents)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                // Write CSV header
                writer.WriteLine("CellName,Formula");

                if (dependents != null)
                {
                    foreach (Cell cell in dependents)
                    {
                        // Escape commas in formula if any
                        string formula = cell.IsFormula ? $"\"{cell.Formula}\"" : "";
                        writer.WriteLine($"{cell.Name},{formula}");
                    }
                }
            }
        }
    }
}