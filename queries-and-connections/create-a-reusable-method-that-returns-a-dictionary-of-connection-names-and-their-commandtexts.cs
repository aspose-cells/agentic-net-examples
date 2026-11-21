// Title: C# method that returns a dictionary of Excel workbook connection names and their CommandText using Aspose.Cells
// Common Searches: how to extract data connection command text from an Excel workbook with Aspose.Cells in C# | Aspose.Cells get list of workbook connections and command texts in .NET | C# retrieve connection names and command texts from an .xlsx file using Aspose.Cells | using reflection to access WorkbookConnections when the property is unavailable in Aspose.Cells | create a dictionary of connection name to command text from an Excel file with Aspose.Cells
// Tags: retrieve workbook connections Aspose.Cells | extract connection command text C# | reflection access WorkbookConnections property | dictionary of connection name to command text | handle missing WorkbookConnections in Aspose.Cells | validate Excel file path before loading

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The GetConnectionCommandTexts method validates the supplied file path, loads the workbook with Aspose.Cells, and uses reflection to safely access the WorkbookConnections collection (if present). It iterates through each connection, reads the Name and CommandText properties, substitutes empty strings for null values, and returns a Dictionary<string, string> mapping each connection name to its command text.
    public static class WorkbookHelper
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>Dictionary where key is connection name and value is command text (empty string if null).</returns>
        public static Dictionary<string, string> GetConnectionCommandTexts(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path must be provided.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' does not exist.", filePath);

            try
            {
                // Load the workbook from the specified file.
                var workbook = new Workbook(filePath);
                var connections = new Dictionary<string, string>();

                // Use reflection to access WorkbookConnections (may not exist in older versions).
                var connProp = workbook.GetType().GetProperty("WorkbookConnections");
                if (connProp != null)
                {
                    var connCollection = connProp.GetValue(workbook) as System.Collections.IEnumerable;
                    if (connCollection != null)
                    {
                        foreach (var connObj in connCollection)
                        {
                            var nameProp = connObj.GetType().GetProperty("Name");
                            var cmdProp = connObj.GetType().GetProperty("CommandText");

                            string name = nameProp?.GetValue(connObj) as string ?? string.Empty;
                            string cmd = cmdProp?.GetValue(connObj) as string ?? string.Empty;

                            connections[name] = cmd ?? string.Empty;
                        }
                    }
                }

                return connections;
            }
            catch (Exception ex)
            {
                // Wrap any exception to provide context.
                throw new InvalidOperationException("Failed to retrieve workbook connections.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect the first argument to be the Excel file path.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsExample <excel-file-path>");
                return;
            }

            string filePath = args[0];

            try
            {
                var connections = WorkbookHelper.GetConnectionCommandTexts(filePath);
                Console.WriteLine($"Found {connections.Count} connection(s) in '{filePath}':");
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
