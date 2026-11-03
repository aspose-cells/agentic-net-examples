// Title: Generate a summary of empty, numeric, string, boolean, date, formula, error and other cell counts in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to count different cell value types in an Excel file with Aspose.Cells C# | Aspose.Cells C# enumerate used range and get statistics for empty and filled cells | C# code to list number of formula, numeric and error cells in a worksheet using Aspose.Cells | retrieve cell type distribution from an .xlsx workbook with Aspose.Cells .NET API | count empty versus populated cells in Excel using Aspose.Cells library
// Tags: cell type counting Aspose.Cells .NET | enumerate used range Aspose.Cells | Excel worksheet cell statistics C# | Aspose.Cells CellValueType classification | formula cell detection Aspose.Cells | empty cell count Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, selects the first worksheet, determines its used range, and iterates through every cell. It uses the IsFormula property to separate formula cells and the CellValueType enum to classify empty, numeric, string, boolean, datetime, error and other cells, incrementing counters for each type. After processing, it prints a concise summary of the counts, while handling missing files and runtime exceptions.
class WorksheetCellSummary
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Counters for each cell type
            int emptyCount = 0;
            int numericCount = 0;
            int stringCount = 0;
            int boolCount = 0;
            int dateTimeCount = 0;
            int formulaCount = 0;
            int errorCount = 0;
            int otherCount = 0;

            // Determine the used range to limit iteration
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Enumerate all cells within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Formula cells are identified via IsFormula property
                    if (cell.IsFormula)
                    {
                        formulaCount++;
                        continue;
                    }

                    // Use the Type property to get the cell's value type
                    switch (cell.Type)
                    {
                        case CellValueType.IsNull:
                            emptyCount++;
                            break;
                        case CellValueType.IsNumeric:
                            numericCount++;
                            break;
                        case CellValueType.IsString:
                            stringCount++;
                            break;
                        case CellValueType.IsBool:
                            boolCount++;
                            break;
                        case CellValueType.IsDateTime:
                            dateTimeCount++;
                            break;
                        case CellValueType.IsError:
                            errorCount++;
                            break;
                        default:
                            otherCount++;
                            break;
                    }
                }
            }

            // Output the summary
            Console.WriteLine("Worksheet Cell Summary:");
            Console.WriteLine($"Empty cells      : {emptyCount}");
            Console.WriteLine($"Numeric cells    : {numericCount}");
            Console.WriteLine($"String cells     : {stringCount}");
            Console.WriteLine($"Boolean cells    : {boolCount}");
            Console.WriteLine($"DateTime cells   : {dateTimeCount}");
            Console.WriteLine($"Formula cells    : {formulaCount}");
            Console.WriteLine($"Error cells      : {errorCount}");
            Console.WriteLine($"Other types      : {otherCount}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
