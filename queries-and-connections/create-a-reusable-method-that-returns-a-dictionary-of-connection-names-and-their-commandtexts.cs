// Title: C# helper to retrieve external connection names and command texts from an Aspose.Cells workbook
// Description: A static method that loads an Excel file with Aspose.Cells, iterates the ExternalConnectionCollection, and returns a case‑insensitive Dictionary where each key is the connection name and each value is its Command text (empty when null). Includes file‑existence validation and can be used in any .NET project.
// Keywords: Aspose.Cells | ExternalConnectionCollection | C# get connection command | Excel data connections | Workbook DataConnections | dictionary of connection names | retrieve command text | Aspose.Cells example
// Common Searches: Aspose.Cells list external connections C# | Get command text from Excel data connection using Aspose | C# dictionary of workbook connection names | How to read DataConnections in Aspose.Cells | Extract external connection command from .xlsx
// Developer Intent: Obtain a dictionary that maps every external data‑connection name in a workbook to its associated Command text.
// Use Cases: Audit all data connections before publishing a workbook to ensure correct queries. | Validate that required connections contain non‑empty command strings as part of a data‑integrity check. | Generate a compliance report listing connection names and their commands for regulatory review. | Log connection details during automated processing of Excel files.
// AI Prompts: Create unit tests for GetConnectionCommands using a mocked Workbook with multiple external connections. | Refactor GetConnectionCommands to accept a Workbook instance instead of a file path and add argument validation. | Extend the method to return a custom object that includes connection name, command text, and connection type (ODBC, OLEDB, etc.).

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// A static method that loads an Excel file with Aspose.Cells, iterates the ExternalConnectionCollection, and returns a case‑insensitive Dictionary where each key is the connection name and each value is its Command text (empty when null). Includes file‑existence validation and can be used in any .NET project.
public static class ConnectionHelper
{
    // Returns a dictionary where the key is the connection name and the value is its Command text.
    public static Dictionary<string, string> GetConnectionCommands(string workbookPath)
    {
        // Verify that the workbook file exists to avoid FileNotFoundException.
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

        // Load the workbook from the specified file.
        Workbook workbook = new Workbook(workbookPath);

        // Get the collection of external connections.
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Prepare the result dictionary.
        var commands = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Iterate through each connection and store its Name and Command.
        foreach (ExternalConnection conn in connections)
        {
            // Some connections may have a null Command; treat it as an empty string.
            string commandText = conn.Command ?? string.Empty;
            commands[conn.Name] = commandText;
        }

        return commands;
    }
}

public class Program
{
    // Entry point of the console application.
    public static void Main(string[] args)
    {
        try
        {
            // Expect the workbook path as the first argument.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <workbookPath>");
                return;
            }

            string workbookPath = args[0];

            // Retrieve connection commands.
            var commands = ConnectionHelper.GetConnectionCommands(workbookPath);

            // Output the results.
            Console.WriteLine($"External connections in workbook: {workbookPath}");
            foreach (var kvp in commands)
            {
                Console.WriteLine($"Name: {kvp.Key}, Command: {kvp.Value}");
            }
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
