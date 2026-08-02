// Title: Export Excel Query Table to CSV with Headers and Data Types – Aspose.Cells C#
// Description: Loads a workbook containing a query table, uses ExportTableOptions (ExportColumnName = true, CheckMixedValueType = true) to export the used range to a DataTable, then writes the DataTable to a CSV file with column headers and proper CSV escaping, preserving original data types.
// Keywords: Aspose.Cells export query table to CSV | C# export Excel to CSV with headers | ExportTableOptions CheckMixedValueType | ExportDataTable Aspose.Cells | CSV escaping C# DataTable | preserve data types when exporting Excel
// Common Searches: how to export a query table to csv using aspose.cells | aspose.cells export excel range with column names | c# write datatable to csv with proper escaping | preserve numeric types when converting excel to csv | export query table from xlsx to csv asp.net
// Developer Intent: Generate a CSV file from an Excel query table while keeping the original column headers and data types intact.
// Use Cases: Create CSV reports from refreshed query tables for downstream analytics. | Supply external systems with exact column names and typed values via CSV. | Automate scheduled batch jobs that convert Excel query results to CSV.
// AI Prompts: Show C# code using Aspose.Cells to export a worksheet range to CSV, preserving headers and handling mixed data types. | Explain the effect of ExportTableOptions.CheckMixedValueType on the DataTable and how numeric values are kept numeric in the CSV output. | Suggest improvements to the CSV writer for culture‑specific number formatting, custom delimiters, and large‑file streaming.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportQueryTableToCsv
{
    // Loads a workbook containing a query table, uses ExportTableOptions (ExportColumnName = true, CheckMixedValueType = true) to export the used range to a DataTable, then writes the DataTable to a CSV file with column headers and proper CSV escaping, preserving original data types.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook containing the query table.
                string sourcePath = "QueryTable.xlsx";

                // Verify that the source file exists before attempting to load it.
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found. Operation aborted.");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(sourcePath);

                // Get the first worksheet (adjust index if needed).
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure export options.
                ExportTableOptions exportOptions = new ExportTableOptions
                {
                    ExportColumnName = true,
                    CheckMixedValueType = true,
                    ExportAsString = false
                };

                // Determine the used range of the worksheet.
                int firstRow = worksheet.Cells.MinRow;
                int firstColumn = worksheet.Cells.MinColumn;
                int totalRows = worksheet.Cells.MaxRow - firstRow + 1;
                int totalColumns = worksheet.Cells.MaxColumn - firstColumn + 1;

                // Export the range to a DataTable.
                DataTable dataTable = worksheet.Cells.ExportDataTable(firstRow, firstColumn, totalRows, totalColumns, exportOptions);

                // Write the DataTable to a CSV file while preserving headers.
                string csvPath = "ExportedQueryTable.csv";
                using (StreamWriter writer = new StreamWriter(csvPath))
                {
                    // Write column headers.
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        writer.Write(EscapeCsvValue(dataTable.Columns[i].ColumnName));
                        if (i < dataTable.Columns.Count - 1) writer.Write(",");
                    }
                    writer.WriteLine();

                    // Write each data row.
                    foreach (DataRow row in dataTable.Rows)
                    {
                        for (int i = 0; i < dataTable.Columns.Count; i++)
                        {
                            object value = row[i];
                            writer.Write(EscapeCsvValue(value?.ToString() ?? string.Empty));
                            if (i < dataTable.Columns.Count - 1) writer.Write(",");
                        }
                        writer.WriteLine();
                    }
                }

                Console.WriteLine($"Query table exported successfully to '{csvPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to escape CSV fields that contain commas, quotes, or line breaks.
        private static string EscapeCsvValue(string value)
        {
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n"))
            {
                // Double up any existing quotes and wrap the field in quotes.
                string escaped = value.Replace("\"", "\"\"");
                return $"\"{escaped}\"";
            }
            return value;
        }
    }
}
