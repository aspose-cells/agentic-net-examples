// Title: Export the first PivotTable from an Excel workbook to a CSV file while keeping headers and original data types using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export pivot table to CSV with column headers | how to preserve numeric and date formats when converting Excel pivot table to CSV in .NET | write pivot table range to CSV file using Aspose.Cells API | Aspose.Cells export pivot table data types to CSV example | convert Excel pivot table to CSV without losing data type information C#
// Tags: Aspose.Cells export pivot table as CSV | retain column headers in CSV output | preserve numeric and date formats in CSV | use TableRange1 to get pivot range | C# write CSV with proper escaping

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an Excel workbook, accesses the first PivotTable on the first worksheet, obtains its displayed range via TableRange1, iterates through each cell, converts values to strings while preserving their original data type representations (string, numeric, boolean, DateTime, error), applies CSV escaping, and writes the rows to a CSV file, ensuring column headers and data type formatting are retained.
class PivotTableToCsvExporter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "pivot_output.csv";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table in the worksheet
            PivotTable pivot = worksheet.PivotTables[0];

            // The displayed area of the pivot table (including headers)
            CellArea tableArea = pivot.TableRange1; // Correct property for pivot range

            // Prepare a StreamWriter for the CSV output
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Iterate through rows of the pivot table range
                for (int row = tableArea.StartRow; row <= tableArea.EndRow; row++)
                {
                    // Build a line for the current row
                    string[] fields = new string[tableArea.EndColumn - tableArea.StartColumn + 1];

                    for (int col = tableArea.StartColumn; col <= tableArea.EndColumn; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];
                        string fieldValue = ConvertCellToString(cell);

                        // Escape double quotes by doubling them and wrap the field in quotes if needed
                        if (fieldValue.Contains(",") || fieldValue.Contains("\"") || fieldValue.Contains("\n"))
                        {
                            fieldValue = $"\"{fieldValue.Replace("\"", "\"\"")}\"";
                        }

                        fields[col - tableArea.StartColumn] = fieldValue;
                    }

                    // Write the CSV line
                    writer.WriteLine(string.Join(",", fields));
                }
            }

            Console.WriteLine($"Pivot table exported to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Convert a cell value to string while preserving its data type representation
    private static string ConvertCellToString(Cell cell)
    {
        if (cell == null || cell.Type == CellValueType.IsNull)
            return string.Empty;

        switch (cell.Type)
        {
            case CellValueType.IsString:
                return cell.StringValue;

            case CellValueType.IsNumeric:
                // Preserve numeric format using invariant culture
                return cell.DoubleValue.ToString(CultureInfo.InvariantCulture);

            case CellValueType.IsBool:
                // Uppercase boolean representation
                return cell.BoolValue ? "TRUE" : "FALSE";

            case CellValueType.IsDateTime:
                // ISO 8601 format for dates/times
                return cell.DateTimeValue.ToString("o", CultureInfo.InvariantCulture);

            case CellValueType.IsError:
                // Use the string representation of the error (e.g., #DIV/0!)
                return cell.StringValue;

            default:
                // Fallback to string representation
                return cell.StringValue;
        }
    }
}
