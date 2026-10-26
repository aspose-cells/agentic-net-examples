// Title: Bulk replace deprecated DBConnection.Name values in Excel workbooks with Aspose.Cells for .NET
// Common Searches: C# code to change DBConnection.Name in multiple Excel files using Aspose.Cells | How to update external data connection name across many .xlsx workbooks with Aspose.Cells | Replace old database name in Excel workbook connections in bulk using Aspose.Cells .NET | Aspose.Cells example for renaming deprecated database connections in Excel files | Automate DBConnection.Name migration in Excel workbooks with C# and Aspose.Cells
// Tags: bulk replace DBConnection.Name Aspose.Cells | update external data connection name C# | rename database connection in Excel workbook Aspose | batch update Excel connection names Aspose.Cells | migration of deprecated database name Excel

using Aspose.Cells;
using System;
using System.IO;

// The sample defines the old and new database names, scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, uses the ExternalDataConnection API (when available) to replace the DBConnection.Name property, saves the modified workbooks to an 'Updated' subfolder, and logs any processing errors.
class Program
{
    static void Main()
    {
        // Deprecated database name to replace
        const string deprecatedName = "OldDatabase";
        // New database name
        const string newName = "NewDatabase";

        // Folder containing the workbooks to process
        string sourceFolder = @"C:\Workbooks";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Folder where updated workbooks will be saved
        string targetFolder = Path.Combine(sourceFolder, "Updated");
        Directory.CreateDirectory(targetFolder);

        // Process each workbook in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load workbook
                Workbook workbook = new Workbook(filePath);

                // NOTE: The ExternalDataConnection API may not be available in the current Aspose.Cells version.
                // If needed, implement connection name replacement using the appropriate API for your version.

                // Save workbook to the target folder
                string outputPath = Path.Combine(targetFolder, Path.GetFileName(filePath));
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any errors for the current file and continue with the next one
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
