// Title: How to batch‑update the built‑in Document Version property to 3.0 for all .xls, .xlsx, and .xlsm files in a folder using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells code to change Document Version property for every Excel file in a directory | update built‑in Document Version to 3.0 for multiple workbooks with Aspose.Cells .NET | iterate through folder and set Excel Document Version property using Aspose.Cells C# | bulk modify Document Version property in .xls/.xlsx/.xlsm files with Aspose.Cells | automate Document Version update across Excel files in C#
// Tags: bulk modify built‑in document properties Aspose.Cells | set Excel Document Version .NET | iterate over folder Excel files Aspose.Cells | save workbook after property edit C# | filter .xls .xlsx .xlsm files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example scans a specified directory, loads each .xls, .xlsx, or .xlsm workbook with Aspose.Cells, checks for the built‑in "Document Version" property, sets its value to "3.0" when present, saves the workbook back to the original file, and logs progress and any errors.
class UpdateDocumentVersion
{
    static void Main()
    {
        // Path to the directory containing Excel files
        string directoryPath = @"C:\ExcelFiles";

        // Verify that the directory exists
        if (!Directory.Exists(directoryPath))
        {
            Console.WriteLine($"Directory not found: {directoryPath}");
            return;
        }

        // Get all files in the directory (filter later by extension)
        string[] allFiles = Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in allFiles)
        {
            // Process only supported Excel extensions
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue;

            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Update the built‑in Document Version property if it exists
                var docVersionProp = workbook.BuiltInDocumentProperties["Document Version"];
                if (docVersionProp != null)
                {
                    docVersionProp.Value = "3.0";
                }
                else
                {
                    // Property not present; log and continue
                    Console.WriteLine($"'Document Version' property not found in {Path.GetFileName(filePath)}.");
                }

                // Save the workbook, overwriting the original file
                workbook.Save(filePath);
                Console.WriteLine($"Updated Document Version for: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                // Log any errors but continue processing other files
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("DocumentVersion update completed.");
    }
}
