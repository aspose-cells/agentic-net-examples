// Title: Export Excel tables to separate CSV files using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, walks through every worksheet and its ListObjects, converts each table's data range to a DataTable, and writes a UTF‑8 CSV file named after the table's DisplayName (or a generated fallback). The code handles header rows and proper CSV escaping.
// Keywords: Aspose.Cells | C# export table to CSV | ListObject CSV export | Excel tables separate CSV files | Workbook table extraction | UTF-8 CSV Aspose | DataTable to CSV | Excel to CSV automation
// Common Searches: Aspose.Cells export each table to CSV | C# export ListObject as separate CSV files | How to save Excel tables as individual CSV files | Export multiple tables from workbook using Aspose.Cells | CSV file naming based on Excel table name
// Developer Intent: Create individual CSV files for every table in an Excel workbook, using the table name for each file.
// Use Cases: Generate per‑table CSV reports for financial or scientific data sets. | Prepare data extracts for migration to databases or data warehouses. | Automate ETL steps that require CSV inputs from Excel tables.
// AI Prompts: Write C# code with Aspose.Cells that exports all ListObjects in a workbook to CSV files, ensuring proper escaping and naming based on DisplayName. | Refactor the export routine to use async file I/O for faster processing of large tables. | Show how to modify the script to export only selected columns from each table while keeping the CSV format.

using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads an Excel workbook, walks through every worksheet and its ListObjects, converts each table's data range to a DataTable, and writes a UTF‑8 CSV file named after the table's DisplayName (or a generated fallback). The code handles header rows and proper CSV escaping.
class ExportTablesToCsv
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through each table (ListObject) in the worksheet
                foreach (ListObject table in worksheet.ListObjects)
                {
                    try
                    {
                        // Export the table's data range to a DataTable
                        DataTable dataTable = table.DataRange.ExportDataTable();

                        // Use DisplayName; if missing, create a fallback name
                        string tableName = !string.IsNullOrEmpty(table.DisplayName)
                            ? table.DisplayName
                            : $"Table_{worksheet.Index}_{table.StartRow}";

                        // Build the CSV file name using the table's name
                        string csvFileName = $"{tableName}.csv";

                        // Write the DataTable content to a CSV file
                        using (StreamWriter writer = new StreamWriter(csvFileName, false, Encoding.UTF8))
                        {
                            // Write column headers
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                writer.Write(dataTable.Columns[col].ColumnName);
                                if (col < dataTable.Columns.Count - 1) writer.Write(",");
                            }
                            writer.WriteLine();

                            // Write each row
                            foreach (DataRow row in dataTable.Rows)
                            {
                                for (int col = 0; col < dataTable.Columns.Count; col++)
                                {
                                    string field = row[col]?.ToString() ?? string.Empty;

                                    // Escape special characters according to CSV rules
                                    if (field.Contains("\"")) field = field.Replace("\"", "\"\"");
                                    if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
                                        field = $"\"{field}\"";

                                    writer.Write(field);
                                    if (col < dataTable.Columns.Count - 1) writer.Write(",");
                                }
                                writer.WriteLine();
                            }
                        }

                        Console.WriteLine($"Table '{tableName}' exported to '{csvFileName}'.");
                    }
                    catch (Exception exTable)
                    {
                        Console.WriteLine($"Failed to export table in worksheet '{worksheet.Name}': {exTable.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
