// Title: C# – Batch set DocumentVersion to 3.0 for all Excel files in a folder with Aspose.Cells
// Description: A console utility that scans a given directory for Excel‑compatible files (xls, xlsx, xlsm, xlsb, ods, csv), loads each workbook with Aspose.Cells, updates the built‑in DocumentVersion property to "3.0", saves the file in place, and logs success or errors.
// Keywords: Aspose.Cells | C# update DocumentVersion | batch modify Excel metadata | set built‑in document properties .NET | process multiple spreadsheets folder | Excel version property automation | document version 3.0
// Common Searches: how to change DocumentVersion for many Excel files using Aspose.Cells | batch update built‑in properties of workbooks C# | set DocumentVersion to 3.0 programmatically | Aspose.Cells iterate folder and modify metadata | C# script to update Excel file properties in bulk
// Developer Intent: Programmatically set the built‑in DocumentVersion property of every workbook in a specified folder to "3.0" using Aspose.Cells for .NET.
// Use Cases: Standardize version metadata across a repository of reports before release. | Prepare a batch of spreadsheets for regulatory compliance with a uniform DocumentVersion. | Integrate metadata correction into a CI/CD pipeline that generates Excel outputs.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate through a directory and set DocumentVersion = "3.0" for each workbook, including robust error handling. | Create a PowerShell wrapper that calls a .NET method to update DocumentVersion for all Excel files in a folder with Aspose.Cells. | Explain how to adapt the sample to modify a custom document property instead of the built‑in DocumentVersion.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A console utility that scans a given directory for Excel‑compatible files (xls, xlsx, xlsm, xlsb, ods, csv), loads each workbook with Aspose.Cells, updates the built‑in DocumentVersion property to "3.0", saves the file in place, and logs success or errors.
    public static class UpdateDocumentVersionInFolder
    {
        // Updates the Built‑in DocumentVersion property of every Excel file in the specified folder to "3.0".
        public static void Run(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder does not exist: {folderPath}");
                return;
            }

            // Define the Excel file extensions to process.
            string[] extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.xlsb", "*.ods", "*.csv" };

            // Collect all matching files.
            var files = new List<string>();
            foreach (var ext in extensions)
            {
                files.AddRange(Directory.GetFiles(folderPath, ext, SearchOption.TopDirectoryOnly));
            }

            foreach (var file in files)
            {
                // Ensure the file still exists before attempting to load.
                if (!File.Exists(file))
                {
                    Console.WriteLine($"File not found (skipped): {Path.GetFileName(file)}");
                    continue;
                }

                try
                {
                    // Load the workbook from the file.
                    Workbook workbook = new Workbook(file);

                    // Set the DocumentVersion built‑in property to "3.0".
                    workbook.BuiltInDocumentProperties.DocumentVersion = "3.0";

                    // Save the workbook back to the same file.
                    workbook.Save(file);

                    Console.WriteLine($"Updated DocumentVersion for: {Path.GetFileName(file)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process {Path.GetFileName(file)}: {ex.Message}");
                }
            }
        }
    }

    // Entry point for the console application.
    public static class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                string folderPath;

                if (args.Length > 0)
                {
                    folderPath = args[0];
                }
                else
                {
                    Console.Write("Enter the folder path containing Excel files: ");
                    folderPath = Console.ReadLine();
                }

                if (string.IsNullOrWhiteSpace(folderPath))
                {
                    Console.WriteLine("No folder path provided. Exiting.");
                    return;
                }

                UpdateDocumentVersionInFolder.Run(folderPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
