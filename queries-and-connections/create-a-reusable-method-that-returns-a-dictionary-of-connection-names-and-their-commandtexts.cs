// Title: Get a dictionary of Excel workbook connection names and their command texts with Aspose.Cells in C#
// AI Prompts: Write a C# method that accepts a file path, loads the workbook with Aspose.Cells, and returns a Dictionary<string,string> where each key is a connection name and each value is its CommandText. | Create a helper that receives an Aspose.Cells Workbook object, uses reflection to iterate over WorkbookConnections, and populates a case‑insensitive dictionary with connection.Name and connection.Command.
// Common Searches: c# aspocells retrieve workbook connection commandtext dictionary | how to list data connections in an .xlsx using Aspose.Cells | extract connection names from Excel file with Aspose.Cells .NET | reflection to access WorkbookConnections property in Aspose.Cells | get connection command strings from Excel workbook programmatically
// Tags: Aspose.Cells workbook connections dictionary | C# extract Excel data connection command text | reflection enumerate WorkbookConnections Aspose.Cells | load Excel file and retrieve connection names .NET | case-insensitive dictionary of connection command texts

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Provides two static helper methods—one that loads an Excel file by path and another that works with an existing Workbook—to return a case‑insensitive Dictionary mapping each workbook connection’s Name to its Command text, using reflection for version‑agnostic access to WorkbookConnections.
    public static class AsposeCellsHelper
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>Dictionary of connection names and their command texts.</returns>
        public static Dictionary<string, string> GetConnectionCommandTexts(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path must be provided.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(filePath);
                return GetConnectionCommandTexts(workbook);
            }
            catch (Exception ex)
            {
                // Wrap any loading errors.
                throw new InvalidOperationException("Failed to load workbook.", ex);
            }
        }

        /// <param name="workbook">An already loaded Aspose.Cells Workbook.</param>
        /// <returns>Dictionary of connection names and their command texts.</returns>
        public static Dictionary<string, string> GetConnectionCommandTexts(Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            var connectionInfo = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Use reflection to stay compatible with different Aspose.Cells versions.
                var connectionsProp = workbook.GetType().GetProperty("WorkbookConnections");
                if (connectionsProp == null) return connectionInfo; // No connections support.

                var connections = connectionsProp.GetValue(workbook) as IEnumerable;
                if (connections == null) return connectionInfo;

                foreach (var conn in connections)
                {
                    if (conn == null) continue;

                    var nameProp = conn.GetType().GetProperty("Name");
                    var commandProp = conn.GetType().GetProperty("Command");

                    string name = nameProp?.GetValue(conn) as string ?? string.Empty;
                    string command = commandProp?.GetValue(conn) as string ?? string.Empty;

                    connectionInfo[name] = command;
                }
            }
            catch (Exception ex)
            {
                // Log or handle reflection errors if needed.
                throw new InvalidOperationException("Failed to retrieve workbook connections.", ex);
            }

            return connectionInfo;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Use first argument as file path or fallback to a default name.
                string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

                var connections = AsposeCellsHelper.GetConnectionCommandTexts(filePath);

                Console.WriteLine($"Connections found in '{filePath}':");
                foreach (var kvp in connections)
                {
                    Console.WriteLine($"Name: {kvp.Key}, CommandText: {kvp.Value}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
