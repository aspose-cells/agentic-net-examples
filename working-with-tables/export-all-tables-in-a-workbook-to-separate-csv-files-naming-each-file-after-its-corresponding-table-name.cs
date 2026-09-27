// Title: Export each Excel table in a workbook to a separate CSV file named after the table’s display name with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx workbook using Aspose.Cells, iterates through all worksheets and their ListObjects, and writes each table to a CSV file whose name is derived from the table’s DisplayName. | Create a helper method that converts an Aspose.Cells ListObject into a DataTable while preserving column headers, then serialize the DataTable to an RFC‑4180 compliant CSV using file‑system‑safe filenames.
// Common Searches: Aspose.Cells C# export all tables in a workbook to individual CSV files | How to save Excel ListObject as CSV with table name using Aspose.Cells .NET | C# loop through worksheets and ListObjects and write each to separate CSV | Generate safe file names from Excel table display names in Aspose.Cells | Convert Aspose.Cells ListObject to DataTable and export to CSV
// Tags: Aspose.Cells export ListObject to CSV | Excel table to CSV conversion .NET | CSV filename from ListObject DisplayName | DataTable generation from Aspose.Cells table | RFC 4180 CSV writing in C#

using System;
using System.IO;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an input.xlsx workbook with Aspose.Cells, iterates over every worksheet and each ListObject (Excel table) it contains, converts each table to a DataTable while preserving column headers, builds a file‑system‑safe CSV filename from the table’s DisplayName, and writes the data to a CSV file using RFC‑4180 escaping, logging each export operation.
class ExportTablesToCsv
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Process each table (ListObject) on the worksheet
                foreach (ListObject table in sheet.ListObjects)
                {
                    try
                    {
                        // Export table data to a DataTable manually (including column names)
                        DataTable dataTable = ExportTableToDataTable(table);

                        // Build a safe CSV file name from the table name
                        string tableName = table.DisplayName ?? "Table";
                        string safeTableName = MakeFileNameSafe(tableName);
                        string csvFilePath = $"{safeTableName}.csv";

                        // Write the DataTable to CSV
                        using (StreamWriter writer = new StreamWriter(csvFilePath))
                        {
                            // Header row
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                writer.Write(EscapeCsvValue(dataTable.Columns[col].ColumnName));
                                if (col < dataTable.Columns.Count - 1)
                                    writer.Write(",");
                            }
                            writer.WriteLine();

                            // Data rows
                            foreach (DataRow row in dataTable.Rows)
                            {
                                for (int col = 0; col < dataTable.Columns.Count; col++)
                                {
                                    writer.Write(EscapeCsvValue(row[col]?.ToString() ?? string.Empty));
                                    if (col < dataTable.Columns.Count - 1)
                                        writer.Write(",");
                                }
                                writer.WriteLine();
                            }
                        }

                        Console.WriteLine($"Exported table '{tableName}' to '{csvFilePath}'.");
                    }
                    catch (Exception exTable)
                    {
                        Console.WriteLine($"Error exporting table '{table.DisplayName}': {exTable.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Convert a ListObject (table) to a DataTable, preserving column headers
    private static DataTable ExportTableToDataTable(ListObject table)
    {
        DataTable dt = new DataTable();

        // Get the range that contains the table data (including headers)
        Aspose.Cells.Range range = table.DataRange;
        int firstRow = range.FirstRow;
        int firstColumn = range.FirstColumn;
        int totalRows = range.RowCount;
        int totalCols = range.ColumnCount;

        // Add columns using the first row as header
        for (int c = 0; c < totalCols; c++)
        {
            string header = range.Worksheet.Cells[firstRow, firstColumn + c].StringValue;
            if (string.IsNullOrEmpty(header))
                header = $"Column{c + 1}";
            // Ensure unique column names
            string uniqueHeader = header;
            int duplicateIndex = 1;
            while (dt.Columns.Contains(uniqueHeader))
            {
                uniqueHeader = $"{header}_{duplicateIndex}";
                duplicateIndex++;
            }
            dt.Columns.Add(uniqueHeader);
        }

        // Add data rows (skip header row)
        for (int r = 1; r < totalRows; r++)
        {
            DataRow dr = dt.NewRow();
            for (int c = 0; c < totalCols; c++)
            {
                var cell = range.Worksheet.Cells[firstRow + r, firstColumn + c];
                dr[c] = cell.Value?.ToString() ?? string.Empty;
            }
            dt.Rows.Add(dr);
        }

        return dt;
    }

    // Escape CSV values according to RFC 4180
    private static string EscapeCsvValue(string value)
    {
        bool mustQuote = value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r");
        if (mustQuote)
        {
            value = value.Replace("\"", "\"\"");
            return $"\"{value}\"";
        }
        return value;
    }

    // Create a file-system‑safe name from the table name
    private static string MakeFileNameSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
