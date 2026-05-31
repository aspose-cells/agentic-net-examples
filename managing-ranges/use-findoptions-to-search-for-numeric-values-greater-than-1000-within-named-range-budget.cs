using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "Budget"
            Aspose.Cells.Range budgetRange = workbook.Worksheets.GetRangeByName("Budget");
            if (budgetRange == null)
            {
                Console.WriteLine("Named range \"Budget\" was not found.");
                return;
            }

            // Build a CellArea that represents the range – this will be used by FindOptions
            CellArea searchArea = new CellArea
            {
                StartRow = budgetRange.FirstRow,
                StartColumn = budgetRange.FirstColumn,
                EndRow = budgetRange.FirstRow + budgetRange.RowCount - 1,
                EndColumn = budgetRange.FirstColumn + budgetRange.ColumnCount - 1
            };

            // Configure FindOptions to limit the search to the named range
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values // search cell values
            };
            findOptions.SetRange(searchArea); // apply the range

            // Collect cells whose numeric value is greater than 1000
            List<Cell> cellsGreaterThanThousand = new List<Cell>();
            Cells cells = workbook.Worksheets[0].Cells;

            // Iterate through the cells inside the defined CellArea
            for (int row = searchArea.StartRow; row <= searchArea.EndRow; row++)
            {
                for (int col = searchArea.StartColumn; col <= searchArea.EndColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Value is double d && d > 1000)
                    {
                        cellsGreaterThanThousand.Add(cell);
                    }
                }
            }

            // Output the results
            Console.WriteLine($"Found {cellsGreaterThanThousand.Count} cells with value > 1000 in the \"Budget\" range:");
            foreach (Cell c in cellsGreaterThanThousand)
            {
                Console.WriteLine($"{c.Name}: {c.Value}");
            }

            // Save the workbook (optional – demonstrates the lifecycle rule)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}