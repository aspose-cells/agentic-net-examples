// Title: Batch Update DBConnection.Name in Multiple Excel Files with Aspose.Cells for .NET
// Description: Scans a folder for .xls and .xlsx workbooks, loads each file with Aspose.Cells, iterates the DataConnections collection, modifies every DBConnection.Name (e.g., adds a suffix), and saves the workbook back to the original location, with error handling and optional logging.
// Keywords: Aspose.Cells | C# | .NET | DBConnection | DataConnections | external data connection | batch rename | Excel workbook update | in‑place save | folder processing | automate Excel connections | update connection name
// Common Searches: Aspose.Cells batch update DBConnection name | C# rename external connections in multiple Excel files | how to modify DataConnections collection with Aspose.Cells | save edited workbook in place using Aspose.Cells | automate Excel DBConnection name change .NET
// Developer Intent: Load each Excel file in a specified directory, change the Name property of all DBConnection objects, and overwrite the original files with the updated workbooks.
// Use Cases: Append version identifiers to all database connection names before releasing reports. | Standardize connection naming across a suite of dashboards after a schema change. | Automate renaming of connections when the underlying database is renamed, eliminating manual edits.
// AI Prompts: Create C# code that prepends a custom prefix to DBConnection.Name for every workbook in a given folder using Aspose.Cells and logs each modification. | Refactor the script to extract the connection‑renaming logic into a reusable method that accepts a delegate for custom naming rules. | Write a robust batch updater that backs up each Excel file before applying the DBConnection name change and provides a summary report.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Scans a folder for .xls and .xlsx workbooks, loads each file with Aspose.Cells, iterates the DataConnections collection, modifies every DBConnection.Name (e.g., adds a suffix), and saves the workbook back to the original location, with error handling and optional logging.
class UpdateDbConnectionNames
{
    static void Main()
    {
        // Directory containing the Excel files
        string folderPath = @"C:\ExcelFiles";

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all .xls and .xlsx files in the directory
        string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");

        foreach (string filePath in excelFiles)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(filePath);

                // Iterate through all external data connections
                foreach (ExternalConnection connection in workbook.DataConnections)
                {
                    // Process only DBConnection objects
                    if (connection is DBConnection dbConn)
                    {
                        // Example update: append "_Updated" to the existing name
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
