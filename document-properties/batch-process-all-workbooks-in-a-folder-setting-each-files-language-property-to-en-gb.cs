// Title: Batch set Excel workbook language to British English (en‑GB) with Aspose.Cells for .NET (C#)
// Description: C# example that scans a folder, loads each Excel‑compatible file with Aspose.Cells, sets the built‑in Document Property Language to “en‑GB”, overwrites the workbook, and logs missing or password‑protected files. Ideal for bulk updating UK locale metadata.
// Keywords: Aspose.Cells C# language property | set Excel language en-GB | batch update document properties Aspose | bulk edit Excel metadata .NET | process multiple workbooks Aspose.Cells | UK locale Excel files | C# Aspose.Cells example | built‑in document properties | Excel language property automation | Aspose.Cells bulk processing
// Common Searches: how to set language property for Excel files using Aspose.Cells | C# batch change workbook language to en-GB | bulk update Excel document properties .NET | Aspose.Cells set built‑in language for multiple workbooks | automate UK locale metadata in Excel files | process folder of Excel files Aspose.Cells C#
// Developer Intent: Set the built‑in Language property of every workbook in a directory to "en-GB" using Aspose.Cells.
// Use Cases: Standardize UK language metadata across corporate spreadsheets before distribution | Prepare a repository of reports for British English compliance by updating the Language property in bulk | Integrate language‑setting step into CI/CD pipelines that generate Excel outputs for UK users | Migrate legacy Excel files to a SharePoint library with consistent en‑GB locale metadata | Ensure regulatory compliance for financial models by enforcing a uniform language setting
// AI Prompts: Write C# code that iterates through a directory and sets Workbook.BuiltInDocumentProperties.Language to "en-GB" using Aspose.Cells, with handling for password‑protected files and error logging. | Modify the batch language‑update script to recurse into subfolders and generate a CSV summary of processed, skipped, and failed files. | Explain how to extend the example to also set Author, Company, and Title properties while preserving the existing Language value. | Create a PowerShell wrapper that calls the compiled C# program to batch update Excel language property on Windows servers. | Provide unit tests for the batch language‑update functionality using NUnit and a mock file system.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

// C# example that scans a folder, loads each Excel‑compatible file with Aspose.Cells, sets the built‑in Document Property Language to “en‑GB”, overwrites the workbook, and logs missing or password‑protected files. Ideal for bulk updating UK locale metadata.
class BatchSetLanguage
{
    static void Main()
    {
        // Path to the folder containing the workbooks
        string folderPath = @"C:\Workbooks";

        // Ensure the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Retrieve all files in the folder (non‑recursive)
        string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Process only Excel‑compatible formats
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".xls" || ext == ".xlsx" || ext == ".xlsm" || ext == ".xlsb" || ext == ".csv")
            {
                // Verify the file still exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (no password supplied)
                    Workbook workbook = new Workbook(filePath);

                    // Set the built‑in document language property to "en-GB"
                    workbook.BuiltInDocumentProperties.Language = "en-GB";

                    // Save (overwrite) the workbook
                    workbook.Save(filePath);
                }
                catch (CellsException ex)
                {
                    // Aspose.Cells throws CellsException for password‑protected files and other issues
                    if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine($"Password‑protected file skipped: {filePath}");
                    }
                    else
                    {
                        Console.WriteLine($"CellsException processing file '{filePath}': {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    // Log any other errors and continue processing remaining files
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}
