// Title: Validate external file existence and replace external link paths in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to verify external Excel file exists before updating links with Aspose.Cells in C# | replace external link path in an Excel workbook after checking file presence using Aspose.Cells | C# Aspose.Cells change external link source file and save workbook to a new location | handle missing API members when updating external links with Aspose.Cells | ensure output directory exists before saving a workbook with Aspose.Cells
// Tags: validate external file existence Aspose.Cells | update external link path workbook Aspose.Cells | dynamic external link handling Aspose.Cells | create output directory before workbook save Aspose.Cells | catch RuntimeBinderException Aspose.Cells external links

using System;
using System.IO;
using Aspose.Cells;

// The example checks that both the target workbook and the new external source file exist, loads the workbook with Aspose.Cells, iterates through its Worksheets.ExternalLinks collection using a dynamic reference, updates each link to the verified new path via ChangeExternalLinkPath, handles version‑specific binder errors, creates the output folder if necessary, and saves the modified workbook to a new file.
class ExternalLinkValidator
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains external links
            string workbookPath = @"C:\Docs\Sample.xlsx";

            // New external file path that should replace the existing link
            string newExternalPath = @"C:\Data\ExternalSource.xlsx";

            // Validate that the workbook file exists
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException("The specified workbook file does not exist.", workbookPath);

            // Validate that the new external file actually exists
            if (!File.Exists(newExternalPath))
                throw new FileNotFoundException("The specified external file does not exist.", newExternalPath);

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all external links in the workbook
            foreach (var linkObj in workbook.Worksheets.ExternalLinks)
            {
                try
                {
                    // Use dynamic to access members that may vary between Aspose.Cells versions
                    dynamic link = linkObj;

                    // Log the current external link path if available
                    string oldPath = link.SourceFileName as string;
                    Console.WriteLine($"Updating external link: {oldPath}");

                    // Change the external link to point to the new, validated path
                    link.ChangeExternalLinkPath(newExternalPath);
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // If the expected members are not present in this version, skip the link
                    Console.WriteLine("Unable to modify external link – API members not found in this version.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing an external link: {ex.Message}");
                }
            }

            // Save the updated workbook
            string outputPath = @"C:\Docs\Sample_Updated.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!Directory.Exists(outputDir) && !string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
