using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UpdateDocumentVersionInFolder
    {
        /// <summary>
        /// Updates the DocumentVersion built‑in property of every Excel file in the specified folder to "3.0".
        /// </summary>
        /// <param name="folderPath">Path to the folder containing Excel files.</param>
        public static void Run(string folderPath)
        {
            // Validate the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Define the Excel file extensions to process
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv" };

            // Enumerate all files with the supported extensions
            foreach (string filePath in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // Skip non‑Excel files

                try
                {
                    // Load the workbook from the file (uses the Workbook(string) constructor)
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Update the built‑in DocumentVersion property
                        workbook.BuiltInDocumentProperties.DocumentVersion = "3.0";

                        // Save the workbook back to the same file (overwrites the original)
                        workbook.Save(filePath);
                    }

                    Console.WriteLine($"Updated DocumentVersion for: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing other files
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }
        }

        // Example entry point
        public static void Main(string[] args)
        {
            // Example usage: pass the target directory as the first argument
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the folder path as an argument.");
                return;
            }

            Run(args[0]);
        }
    }
}