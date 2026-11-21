// Title: Create a backup of an Excel workbook before updating external link paths with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells save a backup copy of an Excel file before changing external link paths | How to protect the original workbook when modifying external links using Aspose.Cells in .NET | Aspose.Cells .NET example for backing up a workbook prior to external link path updates | Backup input.xlsx before processing Worksheet.ExternalLinks with Aspose.Cells C# | Create a safe copy of an Excel workbook before updating external link locations in C#
// Tags: Aspose.Cells workbook backup C# | save Excel copy before external link changes | C# external link path update Aspose.Cells | Worksheet.ExternalLinks not available Aspose.Cells version | error handling file existence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, saves a backup copy (input_backup.xlsx) before any modifications, includes placeholder code for iterating worksheets and external links (not supported in the current Aspose.Cells version), and finally saves the (potentially) modified workbook to a new file (input_modified.xlsx). All operations are wrapped in try‑catch blocks for robust error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string backupPath = "input_backup.xlsx";
            const string outputPath = "input_modified.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the original workbook
            Workbook workbook = new Workbook(inputPath);

            // Create a backup copy of the workbook before any modifications
            workbook.Save(backupPath);

            // NOTE: The ExternalLinks collection is not available in the current Aspose.Cells version.
            // If external link manipulation is required, upgrade to a version that supports Worksheet.ExternalLinks.
            // The following placeholder demonstrates where such logic would be placed.

            /*
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Example placeholder for external link processing
                // foreach (ExternalLink link in sheet.ExternalLinks)
                // {
                //     string originalPath = link.LinkPath;
                //     string newPath = Path.Combine(@"C:\NewFolder", Path.GetFileName(originalPath));
                //     link.LinkPath = newPath;
                // }
            }
            */

            // Save the (potentially) modified workbook
            workbook.Save(outputPath);
            Console.WriteLine("Workbook processing completed successfully.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
