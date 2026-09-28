// Title: Convert absolute external link paths to relative paths in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, enumerates its Worksheets.ExternalLinks collection, and replaces each link's DataSource that is an absolute file path with a path relative to the workbook's folder. | Add console logging that displays the original absolute path and the new relative path for every updated external link, then save the workbook back to the same location.
// Common Searches: how to make external links in an Excel file portable using Aspose.Cells C# | Aspose.Cells C# convert absolute DataSource of external links to relative path | C# Path.GetRelativePath with Aspose.Cells external link collection example | update external link paths in workbook to relative using Aspose.Cells for .NET | save Excel workbook after modifying external link DataSource in C#
// Tags: Aspose.Cells external link relative path conversion | C# calculate relative path for Excel external links | update ExternalLinkCollection DataSource in .NET | make Excel workbook external links portable Aspose.Cells | Path.GetRelativePath usage with Aspose.Cells external links

using System;
using System.IO;
using Aspose.Cells;

namespace UpdateExternalLinks
{
    // The example loads an Excel workbook, iterates through its ExternalLinkCollection, converts any absolute DataSource paths to relative paths based on the workbook's directory using Path.GetRelativePath, updates each link, logs the changes, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains external links
            string workbookPath = @"C:\Data\MainWorkbook.xlsx";

            // Verify the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Base directory for relative path calculation (the workbook's folder)
                string workbookDir = Path.GetDirectoryName(workbookPath);

                // Access the collection of external links
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

                // Iterate through each external link and convert its DataSource to a relative path
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];
                    string originalPath = link.DataSource;

                    // Only process if the path is absolute
                    if (!string.IsNullOrEmpty(originalPath) && Path.IsPathRooted(originalPath))
                    {
                        // Compute relative path from the workbook folder to the external file
                        string relativePath = Path.GetRelativePath(workbookDir, originalPath);

                        // Update the link to use the relative path
                        link.DataSource = relativePath;

                        // Optional: log the change
                        Console.WriteLine($"Updated link: '{originalPath}' -> '{relativePath}'");
                    }
                }

                // Save the workbook (overwrites the original file)
                workbook.Save(workbookPath);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
