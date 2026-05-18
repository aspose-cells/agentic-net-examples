using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace UpdateDbConnectionNames
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the Excel files
            string folderPath = @"C:\ExcelFiles";

            // Get all Excel files (both .xls and .xlsx) in the directory
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");

            foreach (string filePath in excelFiles)
            {
                // Load the workbook from the file (using the standard load rule)
                Workbook workbook = new Workbook(filePath);

                // Iterate through all external connections in the workbook
                foreach (ExternalConnection connection in workbook.DataConnections)
                {
                    // Process only DBConnection objects
                    if (connection is DBConnection dbConn)
                    {
                        // Example: prepend "Updated_" to the existing connection name
                        // Adjust the logic as needed for your scenario
                        dbConn.Name = "Updated_" + dbConn.Name;
                    }
                }

                // Save the workbook back to the same file (in‑place update)
                workbook.Save(filePath);
                Console.WriteLine($"Processed and saved: {Path.GetFileName(filePath)}");
            }

            Console.WriteLine("All files have been processed.");
        }
    }
}