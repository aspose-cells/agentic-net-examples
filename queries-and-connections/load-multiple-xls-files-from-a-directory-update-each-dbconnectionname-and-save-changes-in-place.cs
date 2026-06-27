using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main(string[] args)
    {
        // Directory containing the Excel files. Pass as first argument or set a default path.
        string folderPath = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

        // Get all .xls and .xlsx files in the directory.
        string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");

        foreach (string filePath in excelFiles)
        {
            // Load the workbook from the file.
            Workbook workbook = new Workbook(filePath);

            // Iterate through all external data connections.
            foreach (ExternalConnection connection in workbook.DataConnections)
            {
                // Process only DBConnection objects.
                if (connection is DBConnection dbConn)
                {
                    // Update the connection name as needed.
                    // Example: append "_Updated" to the existing name.
                    dbConn.Name = dbConn.Name + "_Updated";
                }
            }

            // Save the modified workbook back to the same file (in‑place update).
            workbook.Save(filePath);
        }
    }
}