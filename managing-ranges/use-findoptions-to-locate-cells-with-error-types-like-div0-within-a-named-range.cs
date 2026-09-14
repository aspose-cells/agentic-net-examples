// Title: Use Aspose.Cells FindOptions in C# to locate #DIV/0! and other error cells inside a named range
// AI Prompts: Write C# code that leverages Aspose.Cells FindOptions to scan a named range and return the addresses of cells whose value type is IsError. | Show how to apply a red fill style to every error cell identified by FindOptions within a specific named range in an Excel workbook. | Create a reusable method that accepts a Workbook and a range name, uses FindOptions to collect error cells, and returns a List<string> of their cell names.
// Common Searches: Aspose.Cells FindOptions error cells in a named range C# | detect #DIV/0! errors inside a specific range with Aspose.Cells .NET | C# code to highlight Excel error values using Aspose.Cells FindOptions | retrieve addresses of cells containing errors from a named range Aspose.Cells | search for error type cells in an Excel workbook with Aspose.Cells FindOptions
// Tags: Aspose.Cells FindOptions error search | C# locate Excel error cells | named range error detection Aspose.Cells | highlight error values in Excel using Aspose.Cells | retrieve cell addresses with IsError type

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads a workbook, obtains the named range "MyRange", iterates through its cells, checks each cell's Type for CellValueType.IsError, collects the error cells, prints their addresses, and saves the workbook unchanged.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "MyRange"
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("MyRange");
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Collect cells that contain any error (including #DIV/0!) within the named range
            List<Cell> errorCells = new List<Cell>();
            int firstRow = namedRange.FirstRow;
            int firstCol = namedRange.FirstColumn;
            int lastRow = firstRow + namedRange.RowCount - 1;
            int lastCol = firstCol + namedRange.ColumnCount - 1;

            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int col = firstCol; col <= lastCol; col++)
                {
                    Cell cell = namedRange.Worksheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsError)
                    {
                        errorCells.Add(cell);
                    }
                }
            }

            // Output the addresses of the error cells
            foreach (Cell errCell in errorCells)
            {
                Console.WriteLine($"Error found at {errCell.Name}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (no modifications made in this example)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
