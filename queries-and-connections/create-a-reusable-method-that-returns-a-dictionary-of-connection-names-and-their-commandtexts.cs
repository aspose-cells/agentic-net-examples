using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsDemo
{
    public static class WorkbookConnectionHelper
    {
        /// <summary>
        /// Loads a workbook and returns a dictionary of connection names and their command texts.
        /// </summary>
        /// <param name="filePath">Path to the Excel file.</param>
        /// <returns>Dictionary of connection names and command texts.</returns>
        public static Dictionary<string, string> GetConnectionCommands(string filePath)
        {
            var connectionCommands = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                // Prevent FileNotFoundException
                if (!File.Exists(filePath))
                    throw new FileNotFoundException($"File not found: {filePath}");

                // Load the workbook (lifecycle: load)
                var workbook = new Workbook(filePath);

                // Access the collection of external connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Iterate through each connection and capture its Name and Command
                foreach (ExternalConnection conn in connections)
                {
                    if (string.IsNullOrEmpty(conn.Name))
                        continue;

                    // Command property may be null
                    string commandText = conn.Command ?? string.Empty;
                    connectionCommands[conn.Name] = commandText;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error retrieving connections: {ex.Message}");
                // Return whatever has been collected (may be empty)
            }
            return connectionCommands;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Use first argument as file path or fallback to a default name
                string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

                var connections = WorkbookConnectionHelper.GetConnectionCommands(filePath);

                Console.WriteLine($"Found {connections.Count} connection(s) in '{filePath}':");
                foreach (var kvp in connections)
                {
                    Console.WriteLine($"Name: {kvp.Key}, Command: {kvp.Value}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}