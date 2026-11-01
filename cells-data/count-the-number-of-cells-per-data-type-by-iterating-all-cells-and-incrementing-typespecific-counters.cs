// Title: Count each cell data type in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that loops through all worksheets and tallies the number of String, Numeric, Boolean, DateTime, Error, Blank, and Unknown cells. | Create a reusable method that accepts a Workbook object and returns a Dictionary<CellValueType, long> containing the occurrence count for every cell type. | Extend the cell‑type counting example to also compute and display the percentage of each data type relative to the total used cells.
// Common Searches: how to get count of different cell value types in an Excel file using Aspose.Cells C# | Aspose.Cells C# iterate all worksheets and summarize cell data types | C# program to count blank, error and date cells in a workbook with Aspose.Cells | retrieve cell type statistics from Excel workbook using Aspose.Cells .NET
// Tags: Aspose.Cells count cell value types | C# iterate workbook cells Aspose.Cells | cell type statistics Excel .NET | calculate cell data type percentages Aspose.Cells | handle blank and error cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates every used cell across all worksheets, determines each cell's CellValueType, increments counters for string, numeric, boolean, date, error, blank and unknown types, and prints the totals.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Counters for each cell data type
            long stringCount = 0;
            long numericCount = 0;
            long booleanCount = 0;
            long dateCount = 0;
            long errorCount = 0;
            long blankCount = 0;
            long unknownCount = 0;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all used cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Determine the cell's data type and increment the appropriate counter
                    switch (cell.Type)
                    {
                        case CellValueType.IsString:
                            stringCount++;
                            break;
                        case CellValueType.IsNumeric:
                            numericCount++;
                            break;
                        case CellValueType.IsBool:
                            booleanCount++;
                            break;
                        case CellValueType.IsDateTime: // Correct enum member for dates
                            dateCount++;
                            break;
                        case CellValueType.IsError:
                            errorCount++;
                            break;
                        case CellValueType.IsNull:
                            blankCount++;
                            break;
                        default:
                            unknownCount++;
                            break;
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cell type counts:");
            Console.WriteLine($"String : {stringCount}");
            Console.WriteLine($"Numeric: {numericCount}");
            Console.WriteLine($"Boolean: {booleanCount}");
            Console.WriteLine($"Date   : {dateCount}");
            Console.WriteLine($"Error  : {errorCount}");
            Console.WriteLine($"Blank  : {blankCount}");
            Console.WriteLine($"Unknown: {unknownCount}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
