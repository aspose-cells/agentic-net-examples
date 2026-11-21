// Title: Iterate over the used range of an Excel worksheet and print each cell’s address and value using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells loop through used range and display cell addresses | read all cells in an Excel worksheet with MaxDisplayRange using Aspose.Cells | log each cell name and its value while iterating over a worksheet in Aspose.Cells | enumerate cells in the used range of a workbook with Aspose.Cells for .NET | Aspose.Cells example to output cell address and value to console
// Tags: enumerate used range Aspose.Cells | log cell address Aspose.Cells C# | MaxDisplayRange iteration example | worksheet cells enumeration Aspose.Cells | console output Excel cell values Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads an existing Excel file, obtains the worksheet’s MaxDisplayRange to determine the used rows and columns, then iterates through every cell in that range and writes each cell’s name (e.g., A1) and its value to the console, with error handling for missing files or runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the used range of the worksheet
            AsposeRange usedRange = worksheet.Cells.MaxDisplayRange;

            // Determine start and end indices using Range properties
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startColumn = usedRange.FirstColumn;
            int endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Iterate through each cell in the used range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    Cell cell = worksheet.Cells[row, col];
                    // Log the cell address (e.g., "A1") and its value
                    Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
