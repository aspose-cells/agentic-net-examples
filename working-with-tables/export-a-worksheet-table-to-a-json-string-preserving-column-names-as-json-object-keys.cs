// Title: Export a specific Excel worksheet table to a pretty‑printed JSON string with column names as keys using Aspose.Cells for .NET
// AI Prompts: Generate a C# method that opens an Excel workbook, locates a ListObject by its name, and returns the table rows as an indented JSON array where each property name matches the column header. | Write code that validates the workbook, worksheet, and table existence, extracts header names, builds a list of dictionaries for each data row, and serializes the collection to formatted JSON with System.Text.Json.
// Common Searches: how to export an Excel table (ListObject) to JSON using Aspose.Cells in C# | Aspose.Cells convert worksheet table data to JSON string preserving column headers | C# read named table from .xlsx and serialize rows to JSON with System.Text.Json | sample code for exporting a ListObject from an Excel file to a JSON array
// Tags: Aspose.Cells ListObject JSON export | C# serialize Excel table rows | JSON conversion of worksheet table | preserve column headers in JSON output | pretty‑printed JSON from .xlsx using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example defines WorksheetTableExporter.ExportTableToJson, which loads a workbook, verifies the specified worksheet and ListObject, extracts column headers, iterates over the data rows to create dictionaries mapping headers to cell values, and returns an indented JSON string via System.Text.Json.
public class WorksheetTableExporter
{
    /// <param name="filePath">Path to the Excel file.</param>
    /// <param name="sheetName">Name of the worksheet containing the table.</param>
    /// <param name="tableName">Name of the table (ListObject) to export.</param>
    /// <returns>JSON string representing the table data.</returns>
    public static string ExportTableToJson(string filePath, string sheetName, string tableName)
    {
        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Excel file not found: {filePath}");

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Get the target worksheet
        Worksheet worksheet = workbook.Worksheets[sheetName];
        if (worksheet == null)
            throw new ArgumentException($"Worksheet '{sheetName}' not found.");

        // Get the table (ListObject) by name
        ListObject table = worksheet.ListObjects[tableName];
        if (table == null)
            throw new ArgumentException($"Table '{tableName}' not found in worksheet '{sheetName}'.");

        // Retrieve column names from the table header
        List<string> columnNames = new List<string>();
        foreach (ListColumn col in table.ListColumns)
            columnNames.Add(col.Name);

        // Prepare a list to hold each row as a dictionary
        List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();

        // DataRange contains only the data rows (excluding header)
        Aspose.Cells.Range dataRange = table.DataRange;
        int rowCount = dataRange.RowCount;
        int colCount = dataRange.ColumnCount;

        // Iterate over each data row
        for (int r = 0; r < rowCount; r++)
        {
            var rowDict = new Dictionary<string, object>();
            for (int c = 0; c < colCount; c++)
            {
                // Get the cell value
                object cellValue = dataRange[r, c].Value;

                // Use the corresponding column name as the key
                string key = columnNames[c];
                rowDict[key] = cellValue;
            }
            rows.Add(rowDict);
        }

        // Serialize the list of dictionaries to JSON
        var options = new JsonSerializerOptions { WriteIndented = true };
        return JsonSerializer.Serialize(rows, options);
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Example usage
            string filePath = "sample.xlsx";
            string sheetName = "Sheet1";
            string tableName = "Table1";

            string json = WorksheetTableExporter.ExportTableToJson(filePath, sheetName, tableName);
            Console.WriteLine(json);
        }
        catch (Exception ex)
        {
            // Runtime safety: report any errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
