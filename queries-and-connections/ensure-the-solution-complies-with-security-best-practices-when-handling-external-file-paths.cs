// Title: How to securely load and save Excel workbooks with path validation using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells validate external file path before opening workbook | Prevent directory traversal attacks when saving Excel files with Aspose.Cells | Restrict Aspose.Cells workbook read/write to a specific folder in .NET | Create missing directories automatically when saving workbook using Aspose.Cells | Handle unauthorized access exception for out‑of‑scope paths in Aspose.Cells
// Tags: Aspose.Cells secure path validation | C# workbook load with directory whitelist | Aspose.Cells directory traversal mitigation | save Excel workbook to allowed directory .NET | automatic directory creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace SecureAsposeCellsDemo
{
    // Demonstrates how to enforce a whitelist base directory, validate external paths, and safely load, create, and save Aspose.Cells Workbook objects in C#, including automatic directory creation and proper exception handling for unauthorized or missing paths.
    public static class WorkbookHandler
    {
        // Base directory where all workbook files are allowed.
        private static readonly string AllowedBaseDirectory = Path.GetFullPath(@"C:\SecureWorkbooks\");

        // Validates that the provided path is within the allowed base directory.
        private static string GetSafeFullPath(string userPath)
        {
            if (string.IsNullOrWhiteSpace(userPath))
                throw new ArgumentException("File path cannot be null or empty.", nameof(userPath));

            // Resolve to absolute path.
            string fullPath = Path.GetFullPath(userPath);

            // Ensure the resolved path starts with the allowed base directory.
            if (!fullPath.StartsWith(AllowedBaseDirectory, StringComparison.OrdinalIgnoreCase))
                throw new UnauthorizedAccessException("Access to the specified path is denied.");

            return fullPath;
        }

        // Loads a workbook from an external file after validating the path.
        public static Workbook LoadWorkbook(string externalFilePath)
        {
            string safePath = GetSafeFullPath(externalFilePath);

            // Verify the file exists to avoid FileNotFoundException.
            if (!File.Exists(safePath))
                throw new FileNotFoundException($"The workbook file was not found: {safePath}");

            // Load the workbook using Aspose.Cells.
            return new Workbook(safePath);
        }

        // Creates a new empty workbook.
        public static Workbook CreateWorkbook()
        {
            // Create a new workbook instance.
            return new Workbook();
        }

        // Saves a workbook to a specified location after validating the path.
        public static void SaveWorkbook(Workbook workbook, string destinationPath)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            string safePath = GetSafeFullPath(destinationPath);

            // Ensure the directory exists.
            string directory = Path.GetDirectoryName(safePath);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            // Save the workbook using Aspose.Cells.
            workbook.Save(safePath);
        }

        // Example usage demonstrating secure handling.
        public static void Example()
        {
            try
            {
                // Define paths relative to the allowed base directory.
                string inputPath = Path.Combine(AllowedBaseDirectory, "Input", "sample.xlsx");
                string outputPath = Path.Combine(AllowedBaseDirectory, "Output", "result.xlsx");

                Workbook wb;

                // Load an existing workbook securely if it exists; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    wb = LoadWorkbook(inputPath);
                }
                else
                {
                    wb = CreateWorkbook();
                }

                // Perform operations on the workbook (e.g., add a sheet).
                Worksheet sheet = wb.Worksheets[wb.Worksheets.Add()];
                sheet.Cells["A1"].PutValue("Secure Processing");

                // Save the modified workbook securely.
                SaveWorkbook(wb, outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during Example execution: {ex.Message}");
                throw;
            }
        }
    }

    // Entry point for the application.
    public static class Program
    {
        public static void Main()
        {
            try
            {
                WorkbookHandler.Example();
                Console.WriteLine("Workbook processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
