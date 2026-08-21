// Title: Create a migration checklist for Excel external links on network drives with Aspose.Cells for .NET
// Description: A C# console app that loads an Excel workbook, scans its ExternalLinkCollection, identifies links whose DataSource paths are UNC (e.g., \\server\share) or mapped drives (e.g., D:\), and prints a migration checklist with the current path, a network‑drive flag, and recommended actions. The sample also shows how to update the DataSource to a new location and save the workbook.
// Keywords: Aspose.Cells external links | C# detect network drive references | Excel UNC path migration | ExternalLinkCollection Aspose.Cells | update DataSource programmatically | migration checklist Excel | sharepoint onedrive link conversion
// Common Searches: list external workbook links Aspose.Cells .NET | identify UNC paths in Excel using C# | migrate network drive links to cloud storage | change external link datasource with Aspose.Cells | generate checklist for Excel external links
// Developer Intent: Locate every external workbook reference that points to a network drive and produce a step‑by‑step migration checklist, optionally rewriting the links to a new location.
// Use Cases: Audit all network‑based external links before moving source files to SharePoint or OneDrive. | Automate path replacement after copying linked workbooks to a local or cloud folder. | Validate link accessibility and flag entries that require manual intervention.
// AI Prompts: Write C# code using Aspose.Cells to replace every external link DataSource that starts with "\\" with a specified local folder path while preserving the original path. | Show how to export the checklist output to a CSV file instead of the console in the provided program. | Explain safe practices for updating ExternalLink.OriginalDataSource and ExternalLink.DataSource when migrating links to cloud storage.

using System;
using Aspose.Cells;

namespace ExternalLinkMigrationChecklist
{
    // A C# console app that loads an Excel workbook, scans its ExternalLinkCollection, identifies links whose DataSource paths are UNC (e.g., \\server\share) or mapped drives (e.g., D:\), and prints a migration checklist with the current path, a network‑drive flag, and recommended actions. The sample also shows how to update the DataSource to a new location and save the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook that needs to be inspected
            // (Replace \"input.xlsx\" with the actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Prepare a checklist header
            Console.WriteLine("=== External Links Migration Checklist ===");
            Console.WriteLine();

            // Iterate through all external links defined in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];
                string dataSource = link.DataSource; // Current path of the external workbook

                // Identify network drive references:
                //   UNC paths start with "\\" (e.g., \\server\share\file.xlsx)
                //   Mapped drive letters that are typically network drives end with ":\"
                // For simplicity, we treat any path that starts with "\\" or contains ":\"
                bool isNetworkPath = dataSource.StartsWith(@"\\") || dataSource.Contains(@":\");

                if (isNetworkPath)
                {
                    // Output checklist entry for this external link
                    Console.WriteLine($"External Link #{i + 1}");
                    Console.WriteLine($"- Current DataSource : {dataSource}");
                    Console.WriteLine("- Detected as network drive reference.");

                    // Suggested migration actions (customize as needed)
                    Console.WriteLine("- Action Items:");
                    Console.WriteLine("  1. Verify accessibility of the network location.");
                    Console.WriteLine("  2. Copy the external workbook to a local or cloud location if required.");
                    Console.WriteLine("  3. Update the DataSource to the new location using ExternalLink.OriginalDataSource or ExternalLink.DataSource.");
                    Console.WriteLine();

                    // Example of updating the link to a new local path (optional)
                    // string newPath = @"C:\MigratedFiles\" + System.IO.Path.GetFileName(dataSource);
                    // link.OriginalDataSource = newPath; // Preserve original for reference
                    // link.DataSource = newPath;          // Apply the new location
                }
            }

            // If no network links were found, inform the user
            if (externalLinks.Count == 0)
            {
                Console.WriteLine("No external links found in the workbook.");
            }

            // Save the workbook (if any modifications were made above)
            // (Replace \"output.xlsx\" with the desired output file path)
            workbook.Save("output.xlsx");

            Console.WriteLine("Checklist generation completed. Workbook saved as \"output.xlsx\".");
        }
    }
}
