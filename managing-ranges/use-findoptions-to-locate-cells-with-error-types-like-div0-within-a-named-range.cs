// Title: How to detect Excel error values such as #DIV/0! inside a named range using Aspose.Cells for .NET
// Common Searches: C# code to list cells that contain #DIV/0! errors within a named range using Aspose.Cells | Aspose.Cells iterate over a named range and identify error cell types | Find Excel error values in a specific range with Aspose.Cells .NET library | Retrieve addresses of cells with error values from a named range in C# | Detect #N/A and #DIV/0! errors in a workbook using Aspose.Cells
// Tags: error cell detection Aspose.Cells | named range iteration C# | Excel error value search .NET | Aspose.Cells locate #DIV/0! cells | C# workbook error analysis

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, obtains the named range "MyRange", walks through each cell in that range, checks for error values (e.g., #DIV/0!, #N/A), prints the addresses of error cells, and saves the workbook unchanged.
class FindErrorInNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "MyRange"
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("MyRange");
            if (namedRange == null)
            {
                Console.WriteLine("Named range \"MyRange\" not found.");
                return;
            }

            // Iterate through all cells in the named range and detect error cells
            int startRow = namedRange.FirstRow;
            int startColumn = namedRange.FirstColumn;
            int rowCount = namedRange.RowCount;
            int columnCount = namedRange.ColumnCount;

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = workbook.Worksheets[0].Cells[startRow + i, startColumn + j];

                    // Check if the cell contains an error value
                    if (cell.Type == CellValueType.IsError)
                    {
                        // Output the address of the error cell
                        Console.WriteLine($"Error found at {cell.Name}");
                    }
                }
            }

            // Save the workbook (no modifications made in this example)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
