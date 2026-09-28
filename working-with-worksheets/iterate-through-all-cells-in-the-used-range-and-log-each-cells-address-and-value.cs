// Title: Read every cell in the used range of an Excel worksheet and print its address and value with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, obtains the worksheet's MaxDisplayRange, and writes each cell's address and value to the console. | Adjust the iteration to skip cells whose Value is null or an empty string, outputting only populated cells. | Extend the sample to export each cell's address and value to a CSV file instead of console output, using Aspose.Cells.
// Common Searches: C# Aspose.Cells iterate over used range and display cell addresses | How to print each cell value with its address using Aspose.Cells .NET | Aspose.Cells MaxDisplayRange loop example for reading Excel data | Skip empty cells while enumerating worksheet cells with Aspose.Cells | Export cell address and value to CSV using Aspose.Cells in C#
// Tags: iterate used range cells Aspose.Cells C# | print cell address and value Aspose.Cells | export worksheet data to CSV Aspose.Cells | skip empty cells Aspose.Cells iteration | maxdisplayrange enumeration Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads an .xlsx workbook with Aspose.Cells, retrieves the first worksheet's MaxDisplayRange, and iterates through each cell, printing its address (e.g., A1) and value to the console. Includes file existence check and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the used range of the worksheet (returns an Aspose.Cells.Range object)
            Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

            // Determine the start and end indices for rows and columns
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Iterate through each cell in the used range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = worksheet.Cells[row, col];
                    // Output the cell address (e.g., "A1") and its value
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
