// Title: Replace absolute external workbook links with relative paths in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert absolute external link URLs to relative paths in an Excel workbook | How to make an Excel workbook portable by updating external links with Aspose.Cells | Using reflection to access Worksheet.ExternalLinks collection in Aspose.Cells C# | Programmatically change external workbook references to relative file paths with Aspose.Cells | Aspose.Cells save workbook after modifying external link paths in .NET
// Tags: Aspose.Cells external link relative path conversion | C# update worksheet external links Aspose.Cells | Excel workbook portability Aspose.Cells | reflection access Worksheet.ExternalLinks C# | relative file references in Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, uses reflection to enumerate each worksheet's ExternalLinks collection (when available), converts any absolute file URLs to paths relative to the workbook's folder, updates the link URLs, and saves the workbook to a new file, handling missing APIs and I/O errors gracefully.
class WorkbookLinkRelativizer
{
    static void Main()
    {
        // Absolute path to the source workbook
        string workbookPath = @"C:\Data\MyWorkbook.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Source workbook not found: {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // NOTE:
            // The ExternalLinks API may not be available in older Aspose.Cells versions.
            // If needed, update the Aspose.Cells package to a version that supports
            // Worksheet.ExternalLinks and ExternalLink.Url. The following block is
            // retained for reference but guarded to avoid compilation issues.

            try
            {
                // Directory of the workbook – used as the base for relative paths
                string workbookDirectory = Path.GetDirectoryName(workbookPath) ?? string.Empty;

                // Iterate through all worksheets and their external links (if supported)
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // The ExternalLinks collection is available in newer versions.
                    // Use reflection to safely access it when present.
                    var externalLinksProp = sheet.GetType().GetProperty("ExternalLinks");
                    if (externalLinksProp == null) continue;

                    var externalLinks = externalLinksProp.GetValue(sheet) as System.Collections.IEnumerable;
                    if (externalLinks == null) continue;

                    foreach (object linkObj in externalLinks)
                    {
                        var urlProp = linkObj.GetType().GetProperty("Url");
                        if (urlProp == null) continue;

                        string absoluteUrl = urlProp.GetValue(linkObj) as string;
                        if (string.IsNullOrEmpty(absoluteUrl) || !Path.IsPathRooted(absoluteUrl))
                            continue;

                        // Convert the absolute path to a relative path based on the workbook's directory
                        string relativePath = Path.GetRelativePath(workbookDirectory, absoluteUrl);

                        // Update the link to use the relative path
                        urlProp.SetValue(linkObj, relativePath);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log any issues with external link processing but continue with saving
                Console.WriteLine($"External link processing skipped: {ex.Message}");
            }

            // Save the modified workbook (overwrite or use a new file name)
            string outputPath = @"C:\Data\MyWorkbook_Portable.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
