using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsDemo
{
    public static class ConnectionHelper
    {
        /// <summary>
        /// Returns a dictionary where the key is the connection name and the value is its Command text.
        /// </summary>
        /// <param name="workbookPath">Full path to the workbook file.</param>
        /// <returns>Dictionary of connection names and their Command texts.</returns>
        public static Dictionary<string, string> GetConnectionCommands(string workbookPath)
        {
            var result = new Dictionary<string, string>();

            if (!File.Exists(workbookPath))
            {
                Console.Error.WriteLine($"File not found: {workbookPath}");
                return result;
            }

            try
            {
                var workbook = new Workbook(workbookPath);
                return GetConnectionCommands(workbook);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error loading workbook: {ex.Message}");
                return result;
            }
        }

        /// <summary>
        /// Overload that works directly with an existing Workbook instance.
        /// </summary>
        /// <param name="workbook">The Workbook object already loaded or created.</param>
        /// <returns>Dictionary of connection names and their Command texts.</returns>
        public static Dictionary<string, string> GetConnectionCommands(Workbook workbook)
        {
            var dict = new Dictionary<string, string>();

            try
            {
                foreach (ExternalConnection connection in workbook.DataConnections)
                {
                    dict[connection.Name] = connection.Command;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing connections: {ex.Message}");
            }

            return dict;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string workbookPath = args.Length > 0 ? args[0] : "input.xlsx";

            var connections = ConnectionHelper.GetConnectionCommands(workbookPath);

            foreach (var kvp in connections)
            {
                Console.WriteLine($"Name: {kvp.Key}, Command: {kvp.Value}");
            }
        }
    }
}