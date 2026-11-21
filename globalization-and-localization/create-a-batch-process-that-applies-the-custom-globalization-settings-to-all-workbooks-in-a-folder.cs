// Title: Batch apply a custom CultureInfo (e.g., French) to all Excel .xlsx workbooks in a folder using Aspose.Cells for .NET
// Common Searches: C# code to set workbook CultureInfo for every .xlsx file in a directory with Aspose.Cells | batch processing Excel files to change locale to fr-FR using Aspose.Cells .NET | apply custom globalization settings to multiple workbooks programmatically Aspose.Cells | process all Excel workbooks in a folder and set Use1904DateSystem flag with Aspose.Cells | automate applying French culture to Excel files in a folder using Aspose.Cells
// Tags: Aspose.Cells batch CultureInfo update | C# set workbook locale fr-FR | process multiple .xlsx files Aspose.Cells | overwrite Excel workbooks with custom globalization | automated workbook Settings.CultureInfo assignment

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example scans a specified folder for .xlsx files, loads each workbook with Aspose.Cells, assigns a French (fr-FR) CultureInfo to the workbook's Settings, optionally prepares the 1904 date system flag, and saves the workbook back, overwriting the original files.
class GlobalizationBatchProcessor
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string folderPath = @"C:\Workbooks";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Define the custom globalization settings
        CultureInfo customCulture = new CultureInfo("fr-FR");
        bool use1904DateSystem = true; // retained for possible future use

        // Get all Excel files in the folder (top‑level only)
        string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in workbookFiles)
        {
            // Ensure the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Apply custom globalization settings
                workbook.Settings.CultureInfo = customCulture;

                // NOTE: In recent Aspose.Cells versions the 1904 date system is set via
                // WorkbookSettings.Use1904DateSystem property, which may not be available.
                // If needed, uncomment the line below when the property exists.
                // workbook.Settings.Use1904DateSystem = use1904DateSystem;

                // Save the workbook (overwrites the original file)
                workbook.Save(filePath, SaveFormat.Xlsx);
                Console.WriteLine($"Processed: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
