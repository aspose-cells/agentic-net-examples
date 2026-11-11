// Title: Locate cells with values over 1000 in the 'Budget' named range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find cells in a named range where value > 1000 | how to filter numeric values greater than 1000 in Excel named range using Aspose.Cells | retrieve addresses of budget cells exceeding 1000 with Aspose.Cells .NET | search for high budget numbers in a specific named range with Aspose.Cells C# | iterate through a named range and collect cells above a threshold using Aspose.Cells
// Tags: Aspose.Cells numeric threshold search in named range | C# filter cells greater than 1000 Aspose.Cells | retrieve cell addresses from Budget named range Aspose.Cells | iterate over Excel named range using Aspose.Cells C# | search cells by value in .NET Excel workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// The example loads an Excel workbook, accesses the named range "Budget", iterates through each cell, collects those with numeric values greater than 1000, and prints their cell names and values to the console using Aspose.Cells for .NET.
class FindBudgetValues
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "Budget"
            AsposeRange budgetRange = workbook.Worksheets.GetRangeByName("Budget");
            if (budgetRange == null)
            {
                Console.WriteLine("Named range 'Budget' not found.");
                return;
            }

            // List to hold cells that satisfy the condition (> 1000)
            List<Cell> resultCells = new List<Cell>();

            // Iterate through each cell in the named range
            int startRow = budgetRange.FirstRow;
            int endRow = startRow + budgetRange.RowCount - 1;
            int startCol = budgetRange.FirstColumn;
            int endCol = startCol + budgetRange.ColumnCount - 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = budgetRange.Worksheet.Cells[row, col];
                    if (cell.Value is double numericValue && numericValue > 1000)
                    {
                        resultCells.Add(cell);
                    }
                }
            }

            // Output the addresses of cells that meet the criteria
            Console.WriteLine("Cells in 'Budget' with numeric values > 1000:");
            foreach (Cell cell in resultCells)
            {
                Console.WriteLine($"{cell.Name} = {cell.Value}");
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
