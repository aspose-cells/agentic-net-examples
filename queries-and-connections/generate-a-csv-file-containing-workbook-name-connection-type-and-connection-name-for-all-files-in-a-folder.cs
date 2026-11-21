// Title: Generate a CSV listing workbook name, connection type and connection name for every Excel file in a folder using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export workbook connections to CSV for all files in a directory | How to enumerate Excel file connections and write them to a CSV with Aspose.Cells | List connection type and name from multiple .xlsx files using Aspose.Cells in .NET | Aspose.Cells retrieve workbook Connections collection and save as CSV | Batch extract Excel workbook connection metadata with Aspose.Cells C#
// Tags: Aspose.Cells export connections to CSV | enumerate workbook connections C# | batch extract Excel connection metadata | dynamic Connections property Aspose.Cells | process multiple Excel files Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsConnectionExport
{
    // The C# program scans a specified folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, dynamically accesses the Connections collection, extracts each connection's Type and Name, and writes the workbook name, connection type, and connection name to a UTF‑8 CSV file with proper field escaping, handling workbooks without connections and missing files gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel files
            string sourceFolder = @"C:\ExcelFiles";

            // Output CSV file path
            string csvOutputPath = @"C:\Export\WorkbookConnections.csv";

            // Build CSV content
            var csvBuilder = new StringBuilder();

            // Header row
            csvBuilder.AppendLine("WorkbookName,ConnectionType,ConnectionName");

            // Ensure source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Process each Excel file in the folder (top‑level only)
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                // Filter supported Excel extensions
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                    continue;

                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook using Aspose.Cells
                    var workbook = new Workbook(filePath);

                    // Use dynamic to access Connections (may not be available in older versions)
                    IEnumerable<object> connections = GetWorkbookConnections(workbook);

                    // If there are no connections, still write the workbook name with empty fields
                    if (connections == null)
                    {
                        AppendCsvLine(csvBuilder, filePath, string.Empty, string.Empty);
                        continue;
                    }

                    // Iterate through all connections in the workbook
                    foreach (var connObj in connections)
                    {
                        try
                        {
                            dynamic conn = connObj;
                            string connectionType = conn.Type != null ? conn.Type.ToString() : string.Empty;
                            string connectionName = conn.Name != null ? conn.Name.ToString() : string.Empty;

                            AppendCsvLine(csvBuilder, filePath, connectionType, connectionName);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing a connection in '{filePath}': {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(csvOutputPath);
            if (!string.IsNullOrEmpty(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write the CSV content to file
            try
            {
                File.WriteAllText(csvOutputPath, csvBuilder.ToString(), Encoding.UTF8);
                Console.WriteLine($"Export completed: {csvOutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write CSV file: {ex.Message}");
            }
        }

        // Retrieves the Connections collection via dynamic access; returns null if not available
        private static IEnumerable<object>? GetWorkbookConnections(Workbook workbook)
        {
            try
            {
                dynamic dynWorkbook = workbook;
                // Attempt to get the Connections property
                var connections = dynWorkbook.Connections as IEnumerable<object>;
                return connections;
            }
            catch
            {
                // Property not available in this version
                return null;
            }
        }

        // Appends a CSV line with proper escaping
        private static void AppendCsvLine(StringBuilder builder, string filePath, string connectionType, string connectionName)
        {
            string workbookName = Path.GetFileName(filePath);
            workbookName = EscapeCsvField(workbookName);
            connectionType = EscapeCsvField(connectionType);
            connectionName = EscapeCsvField(connectionName);
            builder.AppendLine($"{workbookName},{connectionType},{connectionName}");
        }

        // Escapes CSV fields containing commas, quotes, or line breaks
        private static string EscapeCsvField(string field)
        {
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                field = field.Replace("\"", "\"\"");
                return $"\"{field}\"";
            }
            return field;
        }
    }
}
