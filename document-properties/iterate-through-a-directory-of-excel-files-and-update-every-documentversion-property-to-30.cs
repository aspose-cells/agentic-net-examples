// Title: C# – Batch set DocumentVersion = "3.0" for all Excel files in a folder with Aspose.Cells
// Description: Iterates through a directory, loads each workbook (xls, xlsx, xlsm, ods, csv) using Aspose.Cells, updates the built‑in DocumentVersion property to "3.0", saves the file in its original format, and logs successes or errors.
// Keywords: Aspose.Cells C# | DocumentVersion property | batch update Excel metadata | built‑in document properties .NET | folder enumeration Excel files | set DocumentVersion 3.0 | process multiple workbooks | Excel file automation | C# file iteration | Aspose.Cells example
// Common Searches: Aspose.Cells change DocumentVersion for many workbooks | C# script to update Excel metadata in a folder | batch modify built‑in properties of Excel files | set DocumentVersion to 3.0 using Aspose.Cells .NET | iterate through directory and edit Excel properties
// Developer Intent: Programmatically set the DocumentVersion built‑in property to "3.0" for every Excel workbook in a specified directory.
// Use Cases: Standardize version information across a repository before a product release. | Ensure compliance by applying a uniform DocumentVersion to all generated reports. | Mark a bulk data refresh with a new version number in existing spreadsheets.
// AI Prompts: Create a PowerShell version of the batch DocumentVersion update using Aspose.Cells. | Add detailed logging to the C# script, including file path, previous version, and timestamp. | Show how to verify the DocumentVersion change by reading the property after saving. | Modify the code to skip read‑only files and continue processing the rest of the folder.

using System;
using System.IO;
using Aspose.Cells;

namespace UpdateDocumentVersion
{
    // Iterates through a directory, loads each workbook (xls, xlsx, xlsm, ods, csv) using Aspose.Cells, updates the built‑in DocumentVersion property to "3.0", saves the file in its original format, and logs successes or errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the Excel files
            string folderPath = @"C:\ExcelFiles";

            // Validate directory existence
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Directory not found: {folderPath}");
                return;
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".ods", ".csv" };

            // Iterate through each file with a supported extension
            foreach (string filePath in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue; // Skip non‑Excel files

                try
                {
                    // Load the workbook (uses the Workbook(string) constructor)
                    Workbook workbook = new Workbook(filePath);

                    // Update the built‑in DocumentVersion property to "3.0"
                    workbook.BuiltInDocumentProperties.DocumentVersion = "3.0";

                    // Save the workbook back to the same file (preserves original format)
                    workbook.Save(filePath);

                    Console.WriteLine($"Updated DocumentVersion for: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed.");
        }
    }
}
