// Title: Export a Query Table to CSV with Column Headers and Data Types using Aspose.Cells for .NET (C#)
// Description: C# code that loads an Excel workbook, extracts the query table from the first worksheet, applies ExportTableOptions to retain column names and original .NET data types, and writes the result to a UTF‑8 CSV file with proper field escaping.
// Keywords: Aspose.Cells | C# | .NET | Export query table to CSV | preserve column headers | preserve data types | ExportTableOptions | ExportDataTable | Excel to CSV conversion | UTF-8 CSV export
// Common Searches: how to export a query table from Excel to CSV using Aspose.Cells | Aspose.Cells keep column headers when exporting to CSV | export Excel data types to CSV C# | ExportTableOptions CSV example Aspose.Cells | write DataTable to CSV with proper escaping C#
// Developer Intent: Generate a CSV file from an Excel query table while keeping the original column headers and .NET data types.
// Use Cases: Create CSV reports from refreshed query tables for downstream analytics. | Provide a data feed for web services that requires exact numeric and date formats. | Automate nightly extraction of query tables for archival or compliance purposes.
// AI Prompts: Show C# code that exports a named query table to CSV with column headers and data types using Aspose.Cells. | Explain each ExportTableOptions property and how to tweak them for custom CSV formatting. | Modify the example to export a specific worksheet or a custom range instead of the entire used range.

using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExportQueryTableToCsv
{
    // C# code that loads an Excel workbook, extracts the query table from the first worksheet, applies ExportTableOptions to retain column names and original .NET data types, and writes the result to a UTF‑8 CSV file with proper field escaping.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "QueryTableSource.xlsx";
                const string outputFile = "ExportedQueryTable.csv";

                // Verify that the source workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found. Please ensure the file exists in the application directory.");
                    return;
                }

                // Load the workbook that contains the query table
                Workbook workbook = new Workbook(inputFile);

                // Assume the query table is on the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Configure export options
                ExportTableOptions exportOptions = new ExportTableOptions
                {
                    ExportColumnName = true,
                    ExportAsString = false,
                    CheckMixedValueType = true,
                    FormatStrategy = CellValueFormatStrategy.DisplayStyle
                };

                // Determine the used range of the worksheet (including the query table)
                int firstRow = cells.MinRow;
                int firstColumn = cells.MinColumn;
                int totalRows = cells.MaxRow - cells.MinRow + 1;
                int totalColumns = cells.MaxColumn - cells.MinColumn + 1;

                // Export the range to a DataTable preserving data types
                DataTable dataTable = cells.ExportDataTable(firstRow, firstColumn, totalRows, totalColumns, exportOptions);

                // Write the DataTable to a CSV file
                using (StreamWriter writer = new StreamWriter(outputFile, false, Encoding.UTF8))
                {
                    // Write column headers
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        writer.Write(EscapeCsvValue(dataTable.Columns[i].ColumnName));
                        if (i < dataTable.Columns.Count - 1) writer.Write(",");
                    }
                    writer.WriteLine();

                    // Write rows preserving the original .NET types
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

                Console.WriteLine($"Query table exported to CSV successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper to escape CSV fields containing commas, quotes or newlines
        private static string EscapeCsvValue(string value)
        {
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            {
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }
            return value;
        }
    }
}
