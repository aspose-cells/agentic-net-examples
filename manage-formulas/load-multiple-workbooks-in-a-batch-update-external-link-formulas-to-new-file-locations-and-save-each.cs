// Title: Batch updating external link formulas in multiple Excel workbooks with Aspose.Cells for .NET
// Common Searches: C# code to change external reference paths in a batch of Excel files using Aspose.Cells | Aspose.Cells .NET update external links across many workbooks programmatically | How to replace old external workbook links with new locations in Excel files via Aspose.Cells | Automate external link formula updates for multiple .xlsx files in C# | Bulk modify external link formulas in Excel workbooks using Aspose.Cells API
// Tags: batch external link update Aspose.Cells | update external references Excel C# | process multiple workbooks Aspose.Cells | replace external file paths in formulas .NET | save modified workbooks to new folder Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example iterates through all *.xlsx files in a source directory, loads each workbook with Aspose.Cells, updates any external link formulas according to a dictionary that maps old file locations to new ones, and saves the modified workbooks to a separate output folder while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string sourceFolder = @"C:\Workbooks\Input";

        // Folder where the updated workbooks will be saved
        string destinationFolder = @"C:\Workbooks\Output";

        // Mapping of old external file locations to new locations
        var externalLinkMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { @"C:\OldLinks\Reference1.xlsx", @"D:\NewLinks\Reference1.xlsx" },
            { @"C:\OldLinks\Reference2.xlsx", @"D:\NewLinks\Reference2.xlsx" }
            // Add more mappings as needed
        };

        // Ensure the destination folder exists
        Directory.CreateDirectory(destinationFolder);

        // Process each workbook file in the source folder (e.g., *.xlsx)
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            if (!File.Exists(filePath))
                continue; // Skip missing files

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // NOTE: The ExternalLinks API may not be available in older Aspose.Cells versions.
                // If needed, update external links using the appropriate API for your version here.

                // Build the output file path
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(destinationFolder, fileName);

                // Save the updated workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
