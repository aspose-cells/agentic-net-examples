// Title: C# – Generate CSV of Workbook Connection Types and Names for All Excel Files in a Folder using Aspose.Cells
// Description: Scans a specified directory for *.xlsx files, loads each workbook with Aspose.Cells, extracts every data‑connection’s class name and connection name, and writes a CSV report (Workbook, ConnectionType, ConnectionName) with a header line. Includes folder validation, output path handling, and robust error logging.
// Keywords: Aspose.Cells | C# | .NET | list data connections | export to CSV | Excel workbook connections | batch processing | folder scan | connection type | connection name
// Common Searches: list data connections in multiple Excel files C# | export Excel workbook connections to CSV Aspose.Cells | generate connection inventory for folder of .xlsx files | C# code to read workbook data connections Aspose.Cells | how to get connection type and name from Excel workbook using Aspose.Cells
// Developer Intent: Create a CSV inventory that lists each workbook’s data‑connection type and name for every .xlsx file in a given directory.
// Use Cases: Audit workbooks on a shared drive to identify external data sources before migration. | Provide compliance teams with a connection‑type report to ensure only approved sources are used. | Automate documentation of workbook connections for a data‑governance dashboard. | Generate a quick reference for developers troubleshooting connection‑related errors across many files.
// AI Prompts: Write C# code with Aspose.Cells that scans a folder of .xlsx files and outputs a CSV containing Workbook, ConnectionType, and ConnectionName columns. | Extend the program to also include the connection string in the CSV while keeping the existing columns intact. | Add logging that records files that cannot be opened or processed, then continues with the remaining workbooks.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConnectionReport
{
    // Scans a specified directory for *.xlsx files, loads each workbook with Aspose.Cells, extracts every data‑connection’s class name and connection name, and writes a CSV report (Workbook, ConnectionType, ConnectionName) with a header line. Includes folder validation, output path handling, and robust error logging.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel files
            string folderPath = @"C:\ExcelFiles";

            // Output CSV file path
            string outputCsvPath = @"C:\ExcelFiles\WorkbookConnectionsReport.csv";

            try
            {
                // Verify source folder exists
                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"Source folder does not exist: {folderPath}");
                    return;
                }

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputCsvPath);
                if (string.IsNullOrEmpty(outputDir))
                {
                    Console.WriteLine("Invalid output path.");
                    return;
                }

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a StreamWriter for the CSV file
                using (StreamWriter writer = new StreamWriter(outputCsvPath))
                {
                    // Write CSV header
                    writer.WriteLine("Workbook,ConnectionType,ConnectionName");

                    // Iterate through all Excel files in the folder
                    foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
                    {
                        // Guard against missing files (should not happen, but safe)
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        try
                        {
                            // Load the workbook
                            using (Workbook workbook = new Workbook(filePath))
                            {
                                // Access the collection of data connections
                                var connections = workbook.DataConnections;

                                // Iterate over each connection in the workbook
                                for (int i = 0; i < connections.Count; i++)
                                {
                                    var connection = connections[i];

                                    // Determine the connection type (class name) and its name
                                    string connectionType = connection.GetType().Name;
                                    string connectionName = connection.Name;

                                    // Write a line to the CSV file
                                    writer.WriteLine($"{Path.GetFileName(filePath)},{connectionType},{connectionName}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log workbook‑specific errors and continue processing other files
                            Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                        }
                    }
                }

                Console.WriteLine("Connection report generated at: " + outputCsvPath);
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine("Unexpected error: " + ex.Message);
            }
        }
    }
}
