using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ExportPivotToCsv
{
    static void Main()
    {
        // Path to the workbook that contains the pivot table
        string workbookPath = "input.xlsx";

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(workbookPath);

        // Assume the first worksheet contains the pivot table
        Worksheet sheet = workbook.Worksheets[0];

        // Get the first pivot table in the worksheet
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        PivotTable pivot = sheet.PivotTables[0];

        // Determine the range that covers the whole pivot table (including headers)
        CellArea area = pivot.TableRange1; // range of the pivot table (without page fields)
        int startRow = area.StartRow;
        int startColumn = area.StartColumn;
        int totalRows = area.EndRow - area.StartRow + 1;
        int totalColumns = area.EndColumn - area.StartColumn + 1;

        // Set export options: export column names (headers) and keep original data types
        ExportTableOptions exportOptions = new ExportTableOptions
        {
            ExportColumnName = true,   // first row becomes column names
            ExportAsString = false,    // keep original data types
            CheckMixedValueType = true // handle mixed types safely
        };

        // Export the pivot table range to a DataTable (export rule)
        DataTable dataTable = sheet.Cells.ExportDataTable(startRow, startColumn, totalRows, totalColumns, exportOptions);

        // Path for the resulting CSV file
        string csvPath = "pivot_data.csv";

        // Write the DataTable to CSV while preserving column headers
        using (StreamWriter writer = new StreamWriter(csvPath))
        {
            // Write header line
            string headerLine = string.Join(",", dataTable.Columns.Cast<DataColumn>()
                                            .Select(col => EscapeCsv(col.ColumnName)));
            writer.WriteLine(headerLine);

            // Write each data row
            foreach (DataRow row in dataTable.Rows)
            {
                string line = string.Join(",", dataTable.Columns.Cast<DataColumn>()
                                            .Select(col => EscapeCsv(row[col] == DBNull.Value ? string.Empty : row[col].ToString())));
                writer.WriteLine(line);
            }
        }

        Console.WriteLine($"Pivot table data exported successfully to '{csvPath}'.");
    }

    // Helper method to escape CSV fields according to RFC 4180
    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}