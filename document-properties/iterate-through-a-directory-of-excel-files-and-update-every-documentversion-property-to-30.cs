using System;
using System.IO;
using Aspose.Cells;

namespace UpdateDocumentVersion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the folder containing Excel files
            string folderPath = @"C:\ExcelFiles";

            // Get all Excel files in the folder (including .xls, .xlsx, .xlsm, .ods, etc.)
            string[] excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
            
            foreach (string filePath in excelFiles)
            {
                // Filter only supported Excel extensions
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb" && extension != ".ods")
                    continue;

                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Update the built‑in DocumentVersion property
                    workbook.BuiltInDocumentProperties.DocumentVersion = "3.0";

                    // Save the workbook back to the same file (overwrites original)
                    workbook.Save(filePath);
                    
                    Console.WriteLine($"Updated DocumentVersion for: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}