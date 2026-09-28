// Title: C# example: Update external link paths in an Excel workbook and remove missing‑file links using Aspose.Cells with robust error handling
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates its ExternalLinkCollection, replaces a folder segment in each link’s path, checks if the new file exists, updates the link when the file is present, and removes the link when the file is missing. | Show how to wrap workbook loading, saving, and external‑link updates in try/catch blocks to handle I/O and Aspose.Cells exceptions gracefully. | Demonstrate logging or console output that reports which external links were updated, which were removed, and why, using Aspose.Cells for .NET.
// Common Searches: aspocells c# update external link folder path and delete broken links | how to handle missing external files when updating Excel links with Aspose.Cells | C# check if external link source file exists before assigning DataSource Aspose.Cells | remove invalid external links from workbook using Aspose.Cells .NET | error handling for external link path changes in Excel using Aspose.Cells
// Tags: update external link paths Aspose.Cells | remove broken external links C# | verify external file existence .NET | external link error handling Aspose.Cells | save workbook after external link update

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample loads a workbook, iterates its external links, replaces the old folder segment in each link’s path, verifies that the new file exists, updates the link when the file is found, removes the link when the file is missing, and finally saves the workbook while handling loading, saving, and I/O exceptions.
    public class ExternalLinkPathUpdateWithErrorHandling
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the workbook that contains external links
            string workbookPath = "MainWorkbook.xlsx";

            // Ensure the workbook file exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            Workbook workbook;
            try
            {
                workbook = new Workbook(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate over the external links in reverse order so we can safely remove items
            for (int i = externalLinks.Count - 1; i >= 0; i--)
            {
                ExternalLink link = externalLinks[i];

                // Original data source (e.g., "C:\\OldFolder\\Source.xlsx")
                string originalPath = link.OriginalDataSource;

                // Example of path transformation: replace old folder with new folder
                string updatedPath = originalPath.Replace(@"C:\OldFolder\", @"D:\NewFolder\");

                // Verify that the updated file actually exists
                if (File.Exists(updatedPath))
                {
                    // File exists – update the link to point to the new location
                    link.DataSource = updatedPath;
                    link.OriginalDataSource = updatedPath;
                }
                else
                {
                    // File does not exist – handle the error
                    Console.WriteLine($"External link at index {i} points to a missing file: {updatedPath}");
                    Console.WriteLine("Removing this external link to avoid invalid references.");

                    // Remove the external link; formulas referencing it will be cleared
                    externalLinks.RemoveAt(i);
                }
            }

            // Save the workbook after processing external links
            string outputPath = "MainWorkbook_Updated.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
