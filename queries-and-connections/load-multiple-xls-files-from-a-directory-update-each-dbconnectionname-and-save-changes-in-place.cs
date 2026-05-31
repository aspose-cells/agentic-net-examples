using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class UpdateDbConnectionNames
{
    static void Main()
    {
        // Folder containing the Excel files
        string folderPath = @"C:\Path\To\ExcelFiles";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Retrieve all .xls and .xlsx files in the folder
        string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");

        foreach (string filePath in excelFiles)
        {
            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
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
                        // Example modification: prepend "Updated_" to the existing connection name
                        dbConn.Name = "Updated_" + dbConn.Name;
                    }
                }

                // Save the workbook back to the same file (in‑place update)
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                // Log any errors for the current file and continue with the next one
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("All Excel files have been processed and DBConnection names updated.");
    }
}