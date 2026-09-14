// Title: Rename DBConnection.Name for all external database connections in multiple XLS/XLSX workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a folder, opens each .xls or .xlsx workbook with Aspose.Cells, appends a custom suffix to every DBConnection.Name in the workbook's DataConnections collection, and saves the file back to its original location. | Create a C# utility that reads a JSON mapping of old to new DBConnection names and applies those name changes to all external database connections in every Excel file within a specified directory using Aspose.Cells.
// Common Searches: how to change database connection names in multiple Excel files with Aspose.Cells C# | batch update external DBConnection.Name in .xls and .xlsx using Aspose.Cells | C# program to iterate over DataConnections and rename them in a folder of workbooks | Aspose.Cells rename DBConnection in place for many workbooks | update Excel external connections across many files .NET
// Tags: batch rename DBConnection Aspose.Cells | update external connections in XLSX files C# | in‑place workbook save Aspose.Cells | iterate DataConnections collection .NET | process multiple Excel workbooks folder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace UpdateDbConnectionNames
{
    // A C# console application that enumerates all .xls and .xlsx files in a given directory, loads each workbook with Aspose.Cells, appends "_Updated" (or a custom suffix) to every DBConnection.Name found in the DataConnections collection, and saves the workbook back to the same file, enabling bulk renaming of external database connections.
    class Program
    {
        static void Main()
        {
            // Directory containing the Excel files
            string folderPath = @"C:\ExcelFiles";

            // Verify that the directory exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all Excel files (both .xls and .xlsx) in the directory
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");

            foreach (string filePath in excelFiles)
            {
                try
                {
                    // Ensure the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Iterate through all external connections in the workbook
                    foreach (ExternalConnection connection in workbook.DataConnections)
                    {
                        // Process only DBConnection objects
                        if (connection is DBConnection dbConn)
                        {
                            // Example update: append "_Updated" to the existing connection name
                            dbConn.Name = dbConn.Name + "_Updated";
                        }
                    }

                    // Save the workbook back to the same file (in‑place update)
                    workbook.Save(filePath);
                    Console.WriteLine($"Updated DBConnection names in: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("All files processed.");
        }
    }
}
