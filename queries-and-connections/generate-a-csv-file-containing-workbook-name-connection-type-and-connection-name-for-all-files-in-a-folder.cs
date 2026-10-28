// Title: Generate a CSV listing workbook connection type and name for every Excel file in a folder using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export workbook connections to CSV for all files in a directory | How to list Excel workbook connection types and names with Aspose.Cells .NET | Create CSV of workbook connections from multiple Excel workbooks using Aspose.Cells | Iterate through folder of Excel files and extract connection metadata with Aspose.Cells C#
// Tags: export workbook connections to CSV using Aspose.Cells | enumerate Excel workbook connections .NET | Aspose.Cells workbook connections reflection | C# generate CSV from Excel metadata | process multiple Excel files Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsConnectionExport
{
    // The C# program scans a specified directory for Excel workbooks, loads each file with Aspose.Cells, uses reflection to access the WorkbookConnections collection, extracts each connection's Type and Name, safely escapes CSV fields, and writes rows (WorkbookName, ConnectionType, ConnectionName) to a UTF‑8 CSV file, creating the output folder if necessary.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel files
            string sourceFolder = @"C:\ExcelFiles";

            // Output CSV file path
            string csvOutputPath = @"C:\Export\WorkbookConnections.csv";

            try
            {
                ExportWorkbookConnectionsToCsv(sourceFolder, csvOutputPath);
                Console.WriteLine("Export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during export: {ex.Message}");
            }
        }

        /// <param name="folderPath">Path to the folder containing Excel workbooks.</param>
        /// <param name="csvPath">Path where the CSV file will be saved.</param>
        static void ExportWorkbookConnectionsToCsv(string folderPath, string csvPath)
        {
            // Prepare CSV header
            var sb = new StringBuilder();
            sb.AppendLine("WorkbookName,ConnectionType,ConnectionName");

            // Define supported Excel extensions
            string[] extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.xlsb" };

            // Iterate over each supported file type
            foreach (var ext in extensions)
            {
                foreach (var filePath in Directory.GetFiles(folderPath, ext, SearchOption.TopDirectoryOnly))
                {
                    // Ensure the file exists before loading
                    if (!File.Exists(filePath))
                        continue;

                    try
                    {
                        // Load the workbook using Aspose.Cells
                        using (var workbook = new Workbook(filePath))
                        {
                            // Get the workbook name without path
                            string workbookName = Path.GetFileName(filePath);

                            // Use reflection to access WorkbookConnections (may not exist in older versions)
                            var connectionsProp = workbook.GetType().GetProperty("WorkbookConnections");
                            if (connectionsProp == null)
                                continue; // No connections support in this version

                            var connectionsObj = connectionsProp.GetValue(workbook) as IEnumerable;
                            if (connectionsObj == null)
                                continue;

                            foreach (var conn in connectionsObj)
                            {
                                // Retrieve connection type and name via reflection
                                var typeProp = conn.GetType().GetProperty("Type");
                                var nameProp = conn.GetType().GetProperty("Name");

                                string connectionType = typeProp?.GetValue(conn)?.ToString() ?? string.Empty;
                                string connectionName = nameProp?.GetValue(conn)?.ToString() ?? string.Empty;

                                // Escape commas and other special characters in fields
                                connectionName = EscapeCsvField(connectionName);
                                connectionType = EscapeCsvField(connectionType);
                                workbookName = EscapeCsvField(workbookName);

                                // Append row to CSV
                                sb.AppendLine($"{workbookName},{connectionType},{connectionName}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log and continue with next file
                        Console.Error.WriteLine($"Failed to process '{filePath}': {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(csvPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            // Write all collected data to the CSV file
            File.WriteAllText(csvPath, sb.ToString(), Encoding.UTF8);
        }

        private static string EscapeCsvField(string field)
        {
            if (field.Contains("\""))
                field = field.Replace("\"", "\"\"");

            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
                field = $"\"{field}\"";

            return field;
        }
    }
}
