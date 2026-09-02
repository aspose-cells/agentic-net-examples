// Title: Update the 'TabId' custom document property in every Excel workbook in a folder and export a change log CSV with Aspose.Cells for .NET
// AI Prompts: Generate C# code that enumerates all .xlsx, .xls, and .xlsm files in a given directory, reads the existing 'TabId' custom document property using Aspose.Cells, replaces it with a new GUID (or adds it if missing), saves each workbook, and collects the old and new values. | Create a method that writes a CSV file containing FileName, OldTabId, and NewTabId columns for the processed workbooks, ensuring proper escaping of commas and quotes.
// Common Searches: how to batch replace a custom document property in multiple Excel files using Aspose.Cells C# | c# generate CSV report of changed workbook properties after updating TabId | iterate through a folder of .xlsm files and set a new GUID for a custom property with Aspose.Cells
// Tags: Aspose.Cells batch update custom document property | C# modify Excel workbook TabId GUID | export workbook property changes to CSV | process multiple .xlsx files with Aspose.Cells | custom document property management in Excel using .NET

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace ExcelTabIdUpdater
{
    // Scans a specified folder for Excel workbooks, reads the existing 'TabId' custom document property (if any), replaces or adds it with a new GUID using Aspose.Cells, saves each file, and writes a CSV log that records the file name together with the old and new TabId values, handling CSV escaping.
    class Program
    {
        // Represents a single change record for the summary CSV
        class ChangeRecord
        {
            public string FileName { get; set; }
            public string OldTabId { get; set; }
            public string NewTabId { get; set; }
        }

        static void Main(string[] args)
        {
            // Directory containing the Excel files
            string sourceDirectory = @"C:\ExcelFiles";

            // Ensure the source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory not found: {sourceDirectory}");
                return;
            }

            // Output CSV file path
            string summaryCsvPath = Path.Combine(sourceDirectory, "TabIdChangesSummary.csv");

            // List to hold change records
            List<ChangeRecord> changes = new List<ChangeRecord>();

            // Supported Excel extensions
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsm" };

            // Iterate over each Excel file in the directory
            foreach (string filePath in Directory.GetFiles(sourceDirectory))
            {
                // Skip files that are not Excel workbooks
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue;

                // Verify the file exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Retrieve the existing TabId from custom document properties (if any)
                    string oldTabId = null;
                    if (workbook.CustomDocumentProperties.Contains("TabId"))
                    {
                        oldTabId = workbook.CustomDocumentProperties["TabId"].Value?.ToString();
                    }

                    // Generate a new TabId (using a GUID)
                    string newTabId = Guid.NewGuid().ToString();

                    // Update or add the TabId property
                    if (workbook.CustomDocumentProperties.Contains("TabId"))
                    {
                        workbook.CustomDocumentProperties["TabId"].Value = newTabId;
                    }
                    else
                    {
                        workbook.CustomDocumentProperties.Add("TabId", newTabId);
                    }

                    // Save the workbook back to the same file
                    workbook.Save(filePath);

                    // Record the change for the summary
                    changes.Add(new ChangeRecord
                    {
                        FileName = Path.GetFileName(filePath),
                        OldTabId = oldTabId ?? string.Empty,
                        NewTabId = newTabId
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            try
            {
                // Write the summary CSV
                using (StreamWriter writer = new StreamWriter(summaryCsvPath, false))
                {
                    // Header
                    writer.WriteLine("FileName,OldTabId,NewTabId");

                    // Data rows
                    foreach (var record in changes)
                    {
                        // Escape commas in values if necessary
                        string fileName = EscapeCsv(record.FileName);
                        string oldId = EscapeCsv(record.OldTabId);
                        string newId = EscapeCsv(record.NewTabId);
                        writer.WriteLine($"{fileName},{oldId},{newId}");
                    }
                }

                Console.WriteLine($"Processing complete. Summary written to: {summaryCsvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write summary CSV: {ex.Message}");
            }
        }

        // Helper method to escape CSV fields containing commas or quotes
        private static string EscapeCsv(string field)
        {
            if (field == null)
                return string.Empty;

            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            {
                field = field.Replace("\"", "\"\"");
                return $"\"{field}\"";
            }
            return field;
        }
    }
}
