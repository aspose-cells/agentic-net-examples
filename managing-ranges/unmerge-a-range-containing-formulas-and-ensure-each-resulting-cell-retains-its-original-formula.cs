using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UnmergePreserveFormulas
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the merged range (e.g., B2:D4)
                int firstRow = 1;      // zero‑based index for row 2
                int firstColumn = 1;   // zero‑based index for column B
                int totalRows = 3;     // rows 2‑4
                int totalColumns = 3;  // columns B‑D

                // Put some sample data that the formula will use
                cells[0, 1].PutValue(10); // B1
                cells[0, 2].PutValue(20); // C1
                cells[0, 3].PutValue(30); // D1

                // Set a formula in the top‑left cell of the range
                cells[firstRow, firstColumn].Formula = "=B1+C1+D1";

                // Create the range object and merge the cells
                Aspose.Cells.Range range = worksheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
                range.Merge();

                // Store the formula string before unmerging
                string originalFormula = cells[firstRow, firstColumn].Formula;

                // Unmerge the range – after this only the top‑left cell keeps the formula
                range.UnMerge();

                // Apply the stored formula to every cell that was part of the merged area
                for (int r = firstRow; r < firstRow + totalRows; r++)
                {
                    for (int c = firstColumn; c < firstColumn + totalColumns; c++)
                    {
                        cells[r, c].Formula = originalFormula;
                    }
                }

                // Calculate formulas to populate values
                workbook.CalculateFormula();

                // Save the workbook
                string outputPath = "UnmergedPreserveFormulas.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            UnmergePreserveFormulas.Run();
        }
    }
}