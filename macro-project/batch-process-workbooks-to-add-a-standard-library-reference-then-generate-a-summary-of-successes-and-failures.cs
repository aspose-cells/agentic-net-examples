using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace BatchVbaReferenceAdder
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to process
            string folderPath = @"C:\Workbooks";

            // Standard library reference to add
            const string referenceName = "stdole";
            const string referenceLibId = "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation";

            // Counters for summary
            int successCount = 0;
            int failureCount = 0;
            List<string> failedFiles = new List<string>();

            // Get all macro-enabled Excel files in the folder
            string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsm", SearchOption.TopDirectoryOnly);

            foreach (string filePath in workbookFiles)
            {
                try
                {
                    // Load the workbook (uses the provided load rule)
                    Workbook workbook = new Workbook(filePath);

                    // Ensure the workbook has a VBA project
                    if (workbook.VbaProject != null)
                    {
                        // Add the standard library reference (uses the provided API)
                        workbook.VbaProject.References.AddRegisteredReference(referenceName, referenceLibId);
                    }
                    else
                    {
                        // If there is no VBA project, skip adding the reference
                        Console.WriteLine($"No VBA project found in '{Path.GetFileName(filePath)}'. Skipping.");
                    }

                    // Save the workbook (overwrites the original file using the provided save rule)
                    workbook.Save(filePath);

                    successCount++;
                }
                catch (Exception ex)
                {
                    // Record failure details
                    failureCount++;
                    failedFiles.Add($"{Path.GetFileName(filePath)}: {ex.Message}");
                }
            }

            // Output summary
            Console.WriteLine("Batch processing completed.");
            Console.WriteLine($"Total files processed: {workbookFiles.Length}");
            Console.WriteLine($"Successful updates: {successCount}");
            Console.WriteLine($"Failed updates: {failureCount}");

            if (failedFiles.Count > 0)
            {
                Console.WriteLine("Failed files:");
                foreach (string info in failedFiles)
                {
                    Console.WriteLine($" - {info}");
                }
            }
        }
    }
}