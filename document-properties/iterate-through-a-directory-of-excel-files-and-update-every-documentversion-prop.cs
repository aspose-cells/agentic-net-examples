using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UpdateDocumentVersionInFolder
    {
        public static void Run()
        {
            // Path to the folder containing Excel files
            string folderPath = @"C:\ExcelFiles";

            // Validate folder existence
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Define the file extensions Aspose.Cells can load
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv", ".tsv" };

            // Iterate through each file with a supported extension
            foreach (string filePath in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue; // Skip unsupported files

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Update the built‑in DocumentVersion property
                    workbook.BuiltInDocumentProperties.DocumentVersion = "3.0";

                    // Save the workbook back to the same file
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

    public class Program
    {
        public static void Main(string[] args)
        {
            UpdateDocumentVersionInFolder.Run();
        }
    }
}