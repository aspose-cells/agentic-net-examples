// Title: Aspose.Cells C# – Batch remove macros and protect all worksheets to block printing in SharePoint libraries
// Description: A C# console utility that iterates through Excel files in a SharePoint (or local) folder, deletes any VBA macros, applies worksheet protection with a password that disables printing, and saves the updated workbooks back to the source location using the Aspose.Cells API.
// Keywords: Aspose.Cells batch protect worksheets | C# remove Excel macros | prevent printing Excel programmatically | SharePoint Excel file processing | Aspose.Cells API worksheet protection | automate Excel security C# | bulk Excel macro removal
// Common Searches: how to block printing for all sheets with Aspose.Cells | C# code to delete macros and protect Excel workbooks in SharePoint | batch update Excel files to disable printing using Aspose | remove VBA macros from multiple Excel files programmatically | protect worksheets from printing with Aspose.Cells .NET
// Developer Intent: Automatically strip VBA macros and apply print‑blocking protection to every worksheet in each Excel workbook stored in a SharePoint document library.
// Use Cases: Sanitize financial reports before external distribution by ensuring they are macro‑free and cannot be printed. | Enforce corporate policy on shared templates so users can view data but are blocked from printing it. | Run a nightly job that secures all newly uploaded Excel files in a SharePoint library for compliance purposes.
// AI Prompts: Write an Aspose.Cells C# script that connects to a SharePoint document library, loads each Excel file, removes any VBA macros, protects all worksheets against printing with a configurable password, and saves the changes back to SharePoint. | Refactor the provided program to use async/await, add CSV logging of processed files, and include error handling for SharePoint authentication failures. | Explain how to modify the Protect method to allow editing while still preventing printing, and how to assign a unique password per worksheet based on the file name.

using System;
using System.IO;
using Aspose.Cells;

namespace SharePointWorkbookProcessing
{
    // A C# console utility that iterates through Excel files in a SharePoint (or local) folder, deletes any VBA macros, applies worksheet protection with a password that disables printing, and saves the updated workbooks back to the source location using the Aspose.Cells API.
    class Program
    {
        static void Main()
        {
            // Local folder containing Excel workbooks (adjust the path as needed)
            string folderPath = @"C:\Workbooks";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in files)
            {
                string extension = Path.GetExtension(filePath);
                if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                    !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
                {
                    continue; // Skip non‑Excel files
                }

                try
                {
                    // Load workbook from file
                    Workbook workbook = new Workbook(filePath);

                    // Remove macros if present
                    if (workbook.HasMacro)
                        workbook.RemoveMacro();

                    // Protect each worksheet to prevent printing
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.Protect(ProtectionType.All, "PrintBlockPassword", null);
                    }

                    // Save the modified workbook (overwrite original)
                    workbook.Save(filePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed and updated: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
