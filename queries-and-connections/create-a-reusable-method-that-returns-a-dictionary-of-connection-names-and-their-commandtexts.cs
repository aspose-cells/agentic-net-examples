// Title: Retrieve a dictionary of external connection names and command texts from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# method that takes a workbook file path, loads the file with Aspose.Cells, iterates its DataConnections collection, and returns a Dictionary<string,string> mapping each connection's Name to its Command text. | Create a reusable helper that checks the existence of the Excel file, catches loading errors, and extracts external connection command queries from the workbook using Aspose.Cells.
// Common Searches: asp.net how to list external data connections and their SQL queries from an .xlsx with Aspose.Cells | c# retrieve DataConnections command text from Excel workbook using Aspose.Cells | method to map connection names to command strings in an Aspose.Cells workbook | extract external connection command from Excel file programmatically in .NET
// Tags: Aspose.Cells external connections extraction | C# map connection name to command | DataConnections collection iteration .NET | Excel workbook path validation Aspose.Cells | handle missing workbook file .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The example defines a static GetConnectionCommands method that validates the workbook path, loads the Excel file in read‑only mode with Aspose.Cells, accesses its DataConnections collection, and builds a Dictionary where each key is the external connection name and each value is the associated Command text, while providing robust exception handling. A sample Program demonstrates invoking the helper and printing the results.
public static class ConnectionHelper
{
    /// <param name="filePath">Path to the workbook file.</param>
    /// <returns>Dictionary of connection names and their Command texts.</returns>
    public static Dictionary<string, string> GetConnectionCommands(string filePath)
    {
        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Workbook file not found: {filePath}");

        var connectionCommands = new Dictionary<string, string>();

        try
        {
            // Load the workbook (read‑only operation)
            Workbook workbook = new Workbook(filePath);

            // Access the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through each connection and capture its Name and Command
            foreach (ExternalConnection conn in connections)
            {
                string commandText = conn.Command ?? string.Empty;
                string connectionName = conn.Name ?? string.Empty;

                // Add to dictionary (overwrite if name already exists)
                connectionCommands[connectionName] = commandText;
            }
        }
        catch (Exception ex)
        {
            // Wrap any exception with a more descriptive message
            throw new InvalidOperationException("Failed to retrieve connection commands from the workbook.", ex);
        }

        return connectionCommands;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Expect the workbook path as the first argument
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: <program> <workbook_path>");
            return;
        }

        string workbookPath = args[0];

        try
        {
            var connections = ConnectionHelper.GetConnectionCommands(workbookPath);
            Console.WriteLine("External Connections:");
            foreach (var kvp in connections)
            {
                Console.WriteLine($"Name: {kvp.Key}, Command: {kvp.Value}");
            }
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine(fnfEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
