// Title: Generate a CSV report of workbook names, external connection types, and connection names from all Excel files in a folder using Aspose.Cells (C#)
// AI Prompts: Write a C# console application that scans a given directory, loads each Excel workbook with Aspose.Cells, and writes the workbook file name, external connection class name, and connection name to a CSV file. | Enhance the program to also retrieve the connection string for each external data connection and add it as an additional column in the CSV output. | Implement robust error handling that logs files which cannot be opened or processed, including exception details, while allowing the batch operation to continue.
// Common Searches: C# Aspose.Cells list external data connections in multiple workbooks and export to CSV | how to batch extract connection type and name from Excel files using Aspose.Cells | create a folder-wide connection report for .xlsx files with Aspose.Cells in C#
// Tags: Aspose.Cells external connection extraction | C# generate CSV from workbook metadata | batch process Excel files Aspose.Cells | list data connections in .xlsx using Aspose.Cells | export workbook connection details to CSV

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsConnectionReport
{
    // The example demonstrates a C# console program that iterates over all supported Excel files in a specified directory, loads each workbook with Aspose.Cells, extracts every external data connection's type and name, and writes the workbook filename, connection type, and connection name as rows in a CSV file.
    class Program
    {
        static void Main()
        {
            // Folder containing the workbook files
            string folderPath = @"C:\Workbooks";

            // Output CSV file path
            string csvPath = Path.Combine(folderPath, "WorkbookConnections.csv");

            // Create the CSV file and write the header
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                writer.WriteLine("WorkbookName,ConnectionType,ConnectionName");

                // Get all Excel files in the folder (add more extensions if needed)
                string[] excelExtensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv" };
                var files = Directory.GetFiles(folderPath)
                                     .Where(f => excelExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

                foreach (string file in files)
                {
                    // Load the workbook using the Aspose.Cells constructor (load rule)
                    Workbook workbook = new Workbook(file);

                    // Iterate through each external data connection in the workbook
                    foreach (ExternalConnection connection in workbook.DataConnections)
                    {
                        // Determine the connection type (class name) and its name property
                        string connectionType = connection.GetType().Name;
                        string connectionName = connection.Name;

                        // Write a line to the CSV file
                        writer.WriteLine($"{Path.GetFileName(file)},{connectionType},{connectionName}");
                    }
                }
            }

            Console.WriteLine($"Connection report generated at: {csvPath}");
        }
    }
}
