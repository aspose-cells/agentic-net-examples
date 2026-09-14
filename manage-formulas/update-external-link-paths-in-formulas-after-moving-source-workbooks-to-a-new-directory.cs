// Title: Update external workbook link paths after moving source files using Aspose.Cells for .NET
// AI Prompts: Load a workbook, iterate its Worksheets.ExternalLinks collection, and replace each Workbook‑type link whose SourceFullName starts with an old folder path with a new folder path, then save the workbook. | Programmatically adjust Excel external references in C# by changing the SourceFullName of external links from a deprecated directory to a new location with Aspose.Cells.
// Common Searches: C# Aspose.Cells change external link folder path in Excel workbook | how to update external workbook references after moving source directory in .NET | replace old directory in external links of an Excel file using Aspose.Cells | programmatically fix broken external links after moving source files in Excel
// Tags: external link path replacement Aspose.Cells | Workbook.Worksheets.ExternalLinks update .NET | adjust workbook source directory in formulas | C# modify external workbook references programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, scans its external links, replaces any Workbook‑type link whose path begins with the old source directory with a new directory path, and saves the updated workbook using Aspose.Cells for .NET.
class UpdateExternalLinks
{
    static void Main()
    {
        try
        {
            // Paths for the workbook and external source directories
            string workbookPath = @"C:\Workbooks\MainWorkbook.xlsx";
            string oldExternalDir = @"C:\OldExternalSources\";
            string newExternalDir = @"D:\NewExternalSources\";

            // Verify that the source workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(workbookPath);

            // Get the collection of external links used in the workbook
            var externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link and update its path if it points to the old directory
            foreach (dynamic link in externalLinks)
            {
                try
                {
                    // Process only workbook‑type external links
                    if (link.LinkType != null && link.LinkType.ToString() == "Workbook")
                    {
                        string currentPath = link.SourceFullName as string;

                        // Check if the path starts with the old directory (case‑insensitive)
                        if (!string.IsNullOrEmpty(currentPath) &&
                            currentPath.StartsWith(oldExternalDir, StringComparison.OrdinalIgnoreCase))
                        {
                            // Build the new path by replacing the old directory part with the new one
                            string relativePart = currentPath.Substring(oldExternalDir.Length);
                            string updatedPath = Path.Combine(newExternalDir, relativePart);

                            // Assign the updated path back to the link
                            link.SourceFullName = updatedPath;
                        }
                    }
                }
                catch (Exception linkEx)
                {
                    Console.WriteLine($"Warning: Unable to process a link – {linkEx.Message}");
                }
            }

            // Prepare output path and ensure its directory exists
            string outputPath = @"C:\Workbooks\MainWorkbook_Updated.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated external link paths (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
