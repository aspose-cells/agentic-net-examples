// Title: Handle missing external files while updating link and connection paths in Aspose.Cells (C#)
// Description: Loads a workbook, iterates through its ExternalLinkCollection and DataConnections, replaces an old base folder with a new one, verifies each new file with File.Exists, removes broken links or clears missing connections, logs warnings, and saves the updated workbook.
// Keywords: Aspose.Cells | C# external links | update link path | missing file handling | data connections | .NET workbook | replace folder path | remove broken links | file existence check
// Common Searches: Aspose.Cells replace external link folder path | C# remove broken external links Aspose.Cells | check file existence before updating Aspose.Cells connections | update data connection source file path Aspose.Cells .NET | handle missing external files in workbook with Aspose.Cells
// Developer Intent: Replace old directory paths for all external links and data connections, automatically discard entries whose target files are unavailable, and save the cleaned workbook.
// Use Cases: Migrate workbooks to a new server location while cleaning up links to files that were not moved. | Audit a batch of spreadsheets to generate a report of missing external resources and remove them. | Prepare workbooks for distribution by stripping references to unavailable source files.
// AI Prompts: Generate C# code using Aspose.Cells that updates external link and connection paths, checks File.Exists, and logs missing files without throwing exceptions. | Explain why iterating ExternalLinkCollection in reverse order is required when removing items, and provide an alternative safe removal technique. | Show how to extend console warnings to write a detailed missing‑file report to a CSV or log file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads a workbook, iterates through its ExternalLinkCollection and DataConnections, replaces an old base folder with a new one, verifies each new file with File.Exists, removes broken links or clears missing connections, logs warnings, and saves the updated workbook.
class UpdateExternalLinks
{
    static void Main()
    {
        // Load the workbook (create/load rule)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Define old and new base directories for path replacement
        string oldBase = @"C:\OldFolder\";
        string newBase = @"D:\NewFolder\";

        // Update external links in the workbook
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
        for (int i = externalLinks.Count - 1; i >= 0; i--)
        {
            ExternalLink link = externalLinks[i];
            string currentPath = link.DataSource;

            if (string.IsNullOrEmpty(currentPath))
                continue;

            // Build the new path
            string newPath = currentPath.Replace(oldBase, newBase);

            // Check if the new file exists before assigning
            if (File.Exists(newPath))
            {
                link.DataSource = newPath; // Update path
            }
            else
            {
                // Handle missing file: log warning and remove the link
                Console.WriteLine($"Warning: External file not found at '{newPath}'. Removing link '{currentPath}'.");
                externalLinks.RemoveAt(i);
            }
        }

        // Update file‑based external connections
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            if (!string.IsNullOrEmpty(connection.SourceFile))
            {
                string updatedSource = connection.SourceFile.Replace(oldBase, newBase);

                if (File.Exists(updatedSource))
                {
                    connection.SourceFile = updatedSource; // Update source file path
                }
                else
                {
                    // Handle missing connection file
                    Console.WriteLine($"Warning: Connection source file not found at '{updatedSource}'. Clearing source.");
                    connection.SourceFile = string.Empty;
                }
            }
        }

        // Save the modified workbook (save rule)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
