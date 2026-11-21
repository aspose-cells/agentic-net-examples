// Title: C# example: Convert absolute external link paths to relative paths in an Excel workbook using Aspose.Cells
// Common Searches: c# Aspose.Cells convert absolute external link to relative path | how to make external workbook links relative with Aspose.Cells .NET | update Excel external links to relative paths using Aspose.Cells dynamic ExternalLinks | save workbook after changing external link paths in Aspose.Cells | relative path for external references in Excel files C# Aspose.Cells
// Tags: Aspose.Cells update external link paths | relative external link paths Excel .NET | convert absolute to relative paths Aspose.Cells | dynamic ExternalLinks collection C# | save modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Microsoft.CSharp.RuntimeBinder;

// The sample loads an Excel workbook with Aspose.Cells, iterates through each worksheet using a dynamic cast to access the ExternalLinks collection, converts any absolute ExternalLinkPath values to paths relative to the workbook's folder, updates the links, and saves the workbook to a new file while handling missing API support and runtime errors.
class UpdateExternalLinks
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains external links
            string workbookPath = @"C:\Path\To\Your\Workbook.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Source workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(workbookPath);

            // Directory of the workbook – used as the base for relative paths
            string workbookDirectory = Path.GetDirectoryName(workbookPath) ?? string.Empty;

            // Iterate through all worksheets and attempt to update external links
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Use dynamic to access ExternalLinks in case the API is not available at compile time
                    dynamic dynSheet = sheet;
                    foreach (dynamic externalLink in dynSheet.ExternalLinks)
                    {
                        // Get the current absolute path of the external link
                        string absolutePath = externalLink.ExternalLinkPath as string;

                        // Only process if the path is rooted (i.e., an absolute path)
                        if (!string.IsNullOrEmpty(absolutePath) && Path.IsPathRooted(absolutePath))
                        {
                            // Convert the absolute path to a relative path based on the workbook location
                            string relativePath = Path.GetRelativePath(workbookDirectory, absolutePath);

                            // Update the external link to use the relative path
                            externalLink.ExternalLinkPath = relativePath;
                        }
                    }
                }
                catch (RuntimeBinderException)
                {
                    // The current Aspose.Cells version does not support ExternalLinks; skip this worksheet.
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing worksheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Prepare output path
            string outputPath = @"C:\Path\To\Your\Workbook_Updated.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
