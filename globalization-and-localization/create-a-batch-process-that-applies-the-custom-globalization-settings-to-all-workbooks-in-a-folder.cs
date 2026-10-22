// Title: How to batch apply a custom CultureInfo (e.g., French) to all .xlsx workbooks in a folder using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set workbook CultureInfo for every file in a directory | batch change Excel workbook locale to fr-FR with Aspose.Cells | process multiple .xlsx files to apply custom globalization settings in .NET | automate Excel globalization settings across a folder of workbooks in C# | apply same CultureInfo to all workbooks in a folder using Aspose.Cells
// Tags: batch set workbook CultureInfo Aspose.Cells | apply custom locale to .xlsx files C# | globalization settings for multiple Excel workbooks | process folder of workbooks Aspose.Cells | load and save workbooks with CultureInfo C#

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example scans a specified folder for *.xlsx files, loads each workbook with Aspose.Cells, assigns a custom CultureInfo (e.g., fr-FR) to the workbook's Settings, and saves the file back, overwriting the original while handling missing files and exceptions.
class GlobalizationBatchProcessor
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string sourceFolder = @"C:\InputFolder";

        // Verify that the source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Get all Excel files in the folder (you can adjust the pattern as needed)
        string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

        // Define the custom globalization settings
        CultureInfo customCulture = new CultureInfo("fr-FR"); // Example: French (France)

        foreach (string filePath in workbookFiles)
        {
            // Ensure the file still exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found, skipping: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(filePath);

                // Apply custom globalization settings
                workbook.Settings.CultureInfo = customCulture;
                // Note: Use1904DateSystem property is not available in the current Aspose.Cells version.

                // Save the workbook (overwrites the original file)
                workbook.Save(filePath, SaveFormat.Xlsx);

                Console.WriteLine($"Processed successfully: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
