// Title: How to count cells by data type (string, numeric, boolean, error, blank) in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# count string cells in an Excel workbook | C# iterate all worksheets to get numeric and boolean cell totals with Aspose.Cells | How to retrieve blank and error cell counts from a .xlsx file using Aspose.Cells | Get cell type statistics (string, numeric, bool, error) in C# with Aspose.Cells | Count cells by CellValueType in a .NET Excel processing application
// Tags: cell type aggregation Aspose.Cells | CellValueType enumeration usage .NET | excel worksheet cell iteration Aspose | blank and error cell tally Aspose.Cells | data type distribution analysis Excel .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, iterates through every worksheet and each used cell, classifies cells using Cell.Type (string, numeric, boolean, error, blank), updates counters for each data type, and prints the totals, while handling missing files and unexpected exceptions.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Counters for each cell data type
            int stringCount = 0;
            int numericCount = 0;
            int booleanCount = 0;
            int dateCount = 0;
            int errorCount = 0;
            int blankCount = 0;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to avoid scanning empty rows/columns
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Classify the cell based on its value type
                        switch (cell.Type)
                        {
                            case CellValueType.IsString:
                                stringCount++;
                                break;
                            case CellValueType.IsNumeric:
                                // Aspose.Cells does not expose a direct IsDateTime property in all versions.
                                // Treat all numeric cells as numeric for counting purposes.
                                numericCount++;
                                break;
                            case CellValueType.IsBool:
                                booleanCount++;
                                break;
                            case CellValueType.IsError:
                                errorCount++;
                                break;
                            case CellValueType.IsNull:
                                blankCount++;
                                break;
                            default:
                                // Treat any other type as blank
                                blankCount++;
                                break;
                        }
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cell type counts:");
            Console.WriteLine($"String : {stringCount}");
            Console.WriteLine($"Numeric: {numericCount}");
            Console.WriteLine($"Date   : {dateCount}");
            Console.WriteLine($"Boolean: {booleanCount}");
            Console.WriteLine($"Error  : {errorCount}");
            Console.WriteLine($"Blank  : {blankCount}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
