// Title: C# – Export Pivot Table to CSV with Headers Using Aspose.Cells
// Description: Loads an Excel workbook, finds the first worksheet containing a pivot table, extracts the pivot range (including column names) into a DataTable, and writes it to a CSV file while preserving data types and correctly escaping commas and quotes.
// Keywords: Aspose.Cells export pivot table CSV | C# export pivot table to CSV | pivot table CSV with headers | ExportDataTable Aspose.Cells | Excel pivot to CSV C# | preserve data types CSV export | Aspose.Cells example GitHub
// Common Searches: how to export a pivot table to csv using aspose.cells | c# export excel pivot table with column headers | aspose.cells ExportDataTable pivot example | save pivot table as csv file c# | csv export preserving data types asp.net
// Developer Intent: Generate a CSV file from the first pivot table in an Excel workbook, keeping column headers and original data types intact.
// Use Cases: Create CSV reports from pivot summaries for downstream analytics pipelines. | Provide a data feed for BI tools that require flat‑file input. | Automate periodic email attachments containing pivot‑derived metrics.
// AI Prompts: Write C# code with Aspose.Cells that exports a specific pivot table to CSV, ensuring headers are included and values are escaped. | Show how to modify the example to export only visible rows and use a semicolon delimiter. | Explain how to retain numeric formatting (e.g., currency, dates) when converting pivot table values to CSV with Aspose.Cells.

using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel workbook, finds the first worksheet containing a pivot table, extracts the pivot range (including column names) into a DataTable, and writes it to a CSV file while preserving data types and correctly escaping commas and quotes.
class ExportPivotToCsv
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Locate the first worksheet that contains a pivot table
            Worksheet sheet = null;
            PivotTable pivot = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.PivotTables.Count > 0)
                {
                    sheet = ws;
                    pivot = ws.PivotTables[0];
                    break;
                }
            }

            if (pivot == null || sheet == null)
            {
                Console.WriteLine("No pivot tables found in the workbook.");
                return;
            }

            // Determine the range of the pivot table (excluding page fields)
            CellArea range = pivot.TableRange1;
            int firstRow = range.StartRow;
            int firstColumn = range.StartColumn;
            int totalRows = range.EndRow - range.StartRow + 1;
            int totalColumns = range.EndColumn - range.StartColumn + 1;

            // Export options: include column headers
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportColumnName = true
            };

            // Export pivot range to a DataTable
            DataTable dt = sheet.Cells.ExportDataTable(firstRow, firstColumn, totalRows, totalColumns, exportOptions);

            // Write DataTable to CSV
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Header row
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    writer.Write(dt.Columns[i].ColumnName);
                    if (i < dt.Columns.Count - 1) writer.Write(",");
                }
                writer.WriteLine();

                // Data rows
                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        object value = row[i];
                        string text = value == null || value == DBNull.Value
                            ? string.Empty
                            : Convert.ToString(value, CultureInfo.InvariantCulture);

                        // Escape commas and quotes
                        if (text.Contains(",") || text.Contains("\""))
                        {
                            text = $"\"{text.Replace("\"", "\"\"")}\"";
                        }

                        writer.Write(text);
                        if (i < dt.Columns.Count - 1) writer.Write(",");
                    }
                    writer.WriteLine();
                }
            }

            Console.WriteLine("Pivot table exported to CSV successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
