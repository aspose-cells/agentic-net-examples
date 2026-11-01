// Title: Update DBConnection.Name for each .xls workbook in a folder using Aspose.Cells for .NET
// Common Searches: C# how to rename data connections in multiple .xls files with Aspose.Cells | batch update Excel 97‑2003 workbook data connection names using Aspose.Cells .NET | iterate workbook.DataConnections and save changes in place in C# | process all .xls files in a directory to set DBConnection.Name with Aspose.Cells | Aspose.Cells example for updating DBConnection.Name across many workbooks
// Tags: update DBConnection.Name Aspose.Cells | batch process .xls workbooks C# | iterate workbook.DataConnections Aspose | in‑place save Excel97To2003 Aspose.Cells | rename data connections by file name C#

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a given folder for *.xls files, loads each workbook with Aspose.Cells, loops through its DataConnections, sets each connection's Name to the workbook's file name (without extension), and saves the workbook back to the original file using the Excel97To2003 format, with error handling for missing files and exceptions.
class UpdateDbConnectionNames
{
    static void Main()
    {
        // Directory containing the XLS files
        string folderPath = @"C:\Path\To\XlsFiles";

        // Validate the directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all .xls files in the directory
        string[] xlsFiles = Directory.GetFiles(folderPath, "*.xls", SearchOption.TopDirectoryOnly);

        foreach (string filePath in xlsFiles)
        {
            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook from the file
                Workbook workbook = new Workbook(filePath);

                // Iterate through all data connections in the workbook (if any)
                foreach (var connection in workbook.DataConnections)
                {
                    // Set the connection name to the file name without extension
                    connection.Name = Path.GetFileNameWithoutExtension(filePath);
                }

                // Save the workbook back to the same file (in‑place update)
                workbook.Save(filePath, SaveFormat.Excel97To2003);
                Console.WriteLine($"Updated connections in: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
