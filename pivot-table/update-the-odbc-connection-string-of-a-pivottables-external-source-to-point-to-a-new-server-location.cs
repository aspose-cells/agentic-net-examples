// Title: How to programmatically update the ODBC connection string of Excel PivotTables using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that opens an existing .xlsx workbook, enumerates all PivotTables, detects ODBC data sources, and replaces the Server and Database parameters in each PivotTable’s DataSource connection string. | Write a C# helper method that updates a specific keyword in a connection string and applies it to every PivotTable in a workbook before saving the file.
// Common Searches: aspnet change ODBC server name in Excel pivot table programmatically | c# update pivot table external data source connection string using Aspose.Cells | modify Excel pivot table data source to new database with Aspose.Cells .NET | iterate through worksheets and pivot tables to edit ODBC connection string in C# | Aspose.Cells replace Server and Database values in PivotTable DataSource
// Tags: Aspose.Cells modify PivotTable ODBC connection | C# update Excel pivot data source connection string | programmatic server name replacement in pivot ODBC | iterate pivot tables Aspose.Cells .NET | replace database keyword in ODBC connection string C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, loops through each worksheet and its PivotTables, identifies ODBC data sources, replaces the Server/Data Source and Database/Initial Catalog keywords with new values, assigns the revised connection string back to the PivotTable, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // New ODBC server and database names
            string newServer = "NewServerName";
            string newDatabase = "NewDatabaseName";

            // Iterate through worksheets and their pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    try
                    {
                        // PivotTable.DataSource is a string[] in some Aspose.Cells versions.
                        // Use the first element if available.
                        string connStr = (pivot.DataSource != null && pivot.DataSource.Length > 0)
                            ? pivot.DataSource[0]
                            : string.Empty;

                        // Process only ODBC data sources by checking the connection string
                        if (connStr.IndexOf("ODBC", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            connStr.IndexOf("Provider=MSDASQL", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Update Server / Data Source
                            if (connStr.IndexOf("Server=", StringComparison.OrdinalIgnoreCase) >= 0)
                                connStr = ReplaceKeyword(connStr, "Server", newServer);
                            else if (connStr.IndexOf("Data Source=", StringComparison.OrdinalIgnoreCase) >= 0)
                                connStr = ReplaceKeyword(connStr, "Data Source", newServer);

                            // Update Database / Initial Catalog
                            if (connStr.IndexOf("Database=", StringComparison.OrdinalIgnoreCase) >= 0)
                                connStr = ReplaceKeyword(connStr, "Database", newDatabase);
                            else if (connStr.IndexOf("Initial Catalog=", StringComparison.OrdinalIgnoreCase) >= 0)
                                connStr = ReplaceKeyword(connStr, "Initial Catalog", newDatabase);

                            // Assign the modified connection string back to the pivot table
                            // DataSource expects a string[]; wrap the updated string.
                            pivot.DataSource = new[] { connStr };
                        }
                    }
                    catch (Exception exPivot)
                    {
                        Console.WriteLine($"Failed to update pivot '{pivot.Name}' on sheet '{sheet.Name}': {exPivot.Message}");
                    }
                }
            }

            // Save the updated workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Failed to save workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to replace a keyword's value in a connection string
    private static string ReplaceKeyword(string connectionString, string keyword, string newValue)
    {
        // Split into parts separated by ';'
        string[] parts = connectionString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < parts.Length; i++)
        {
            string[] kv = parts[i].Split(new[] { '=' }, 2);
            if (kv.Length == 2 && kv[0].Trim().Equals(keyword, StringComparison.OrdinalIgnoreCase))
            {
                parts[i] = $"{kv[0]}={newValue}";
                break;
            }
        }
        return string.Join(";", parts) + ";";
    }
}
