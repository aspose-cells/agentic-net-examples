// Title: Replace TODAY() formulas with a static snapshot date in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace TODAY() function with a fixed DATE value in Excel | How to substitute dynamic TODAY() formulas with a snapshot date using Aspose.Cells | Iterate all cells in a workbook and change TODAY() to DATE(2023,10,20) with Aspose.Cells | Aspose.Cells recalculate formulas after modifying TODAY() to static date | Create Excel file with static date instead of TODAY() using Aspose.Cells .NET
// Tags: replace TODAY() formula Aspose.Cells | static DATE literal Excel C# | Aspose.Cells calculate formulas after edit | iterate worksheets cells Aspose.Cells | snapshot date substitution Excel formulas

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook (creating a blank file if missing), defines a snapshot DateTime (2023‑10‑20), iterates through every worksheet and cell, replaces any formula containing the TODAY() function with a DATE(year,month,day) literal based on the snapshot, recalculates all formulas, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists; create a blank workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                    var tempWorkbook = new Workbook();
                    tempWorkbook.Save(inputPath);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Snapshot date to replace TODAY().
                DateTime snapshotDate = new DateTime(2023, 10, 20);
                string dateLiteral = $"DATE({snapshotDate.Year},{snapshotDate.Month},{snapshotDate.Day})";

                // Iterate through all worksheets and cells.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (!string.IsNullOrEmpty(cell.Formula) && cell.Formula.Contains("TODAY()"))
                            {
                                // Replace TODAY() with the static DATE literal.
                                string newFormula = cell.Formula.Replace("TODAY()", dateLiteral);
                                cell.Formula = newFormula;
                            }
                        }
                    }
                }

                // Recalculate all formulas.
                workbook.CalculateFormula();

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
