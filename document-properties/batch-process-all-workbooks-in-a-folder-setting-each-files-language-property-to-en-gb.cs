// Title: C# – Batch set Excel workbooks’ Language property to en‑GB with Aspose.Cells
// Description: Scans a folder for .xls, .xlsx, .xlsm and .xlsb files, loads each workbook with Aspose.Cells, sets the built‑in document property Language to "en-GB", saves the file in place, and automatically skips password‑protected workbooks while logging any errors.
// Keywords: Aspose.Cells | C# | set language property | en-GB | batch update Excel | document properties | process multiple workbooks | skip password protected | Workbook.Save | UK locale
// Common Searches: how to change language property of many Excel files using Aspose.Cells | C# batch update built‑in document property Language for Excel workbooks | skip password protected Excel files when setting document properties | Aspose.Cells set Language=en‑GB for all files in a folder | automate Excel metadata changes with C#
// Developer Intent: Update the Language built‑in document property to en‑GB for every Excel workbook in a given directory, ignoring password‑protected files and handling errors gracefully.
// Use Cases: Standardize UK language metadata across corporate reports before distribution. | Prepare a bulk of workbooks for UK locale compliance during a data migration. | Automate maintenance of document properties while safely bypassing protected files.
// AI Prompts: Write C# code with Aspose.Cells that sets the Language property to en‑GB for all Excel files in a folder, skipping password‑protected workbooks. | Show how to extend the script to write the names of successfully processed files to a CSV log. | Provide a version that recursively processes subfolders while applying the same language update. | Explain how to add custom error handling to retry files that fail due to transient I/O issues.

using System;
using System.IO;
using Aspose.Cells;

// Scans a folder for .xls, .xlsx, .xlsm and .xlsb files, loads each workbook with Aspose.Cells, sets the built‑in document property Language to "en-GB", saves the file in place, and automatically skips password‑protected workbooks while logging any errors.
class BatchSetLanguage
{
    static void Main()
    {
        // Path to the folder containing the workbooks
        string folderPath = @"C:\Workbooks";

        // Retrieve all files in the folder
        string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Ensure the file actually exists
            if (!File.Exists(filePath))
                continue;

            // Process only Excel workbook files
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                continue;

            try
            {
                // Load the workbook (no password supplied). If the file is password‑protected,
                // a CellsException will be thrown and caught below.
                LoadOptions loadOptions = new LoadOptions();
                using (Workbook workbook = new Workbook(filePath, loadOptions))
                {
                    // Set the built‑in document property Language to "en-GB"
                    workbook.BuiltInDocumentProperties.Language = "en-GB";

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath);
                }
            }
            catch (CellsException ex) when (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Skip password‑protected files
                Console.WriteLine($"Skipping password‑protected file: {filePath}");
            }
            catch (Exception ex)
            {
                // Log any other unexpected errors and continue processing remaining files
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
