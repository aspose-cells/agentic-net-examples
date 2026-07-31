// Title: Detect Network External Links in Excel and Generate a Migration Checklist with Aspose.Cells for .NET
// Description: Loads an Excel workbook, scans Worksheets.ExternalLinks for UNC or mapped‑drive paths, prints a migration checklist for each link, updates the DataSource to a new local folder while preserving the original path, recalculates formulas, and saves the modified file.
// Keywords: Aspose.Cells external links | detect UNC paths in Excel | update external data source Aspose.Cells | Excel migration checklist | recalculate formulas after link change | C# external workbook references | network drive Excel links
// Common Searches: list external workbook links with UNC paths using Aspose.Cells | replace network external links with local folder in C# | generate migration checklist for Excel external links | recalculate formulas after updating external links .NET | how to identify mapped‑drive links in an Excel file programmatically
// Developer Intent: Find all external workbook references that point to network locations, produce a step‑by‑step migration checklist, and rewrite those links to a new local directory while keeping the original paths for reference.
// Use Cases: Audit a workbook before moving it to a new server by reporting every network‑based external link. | Automate the migration of external data sources from UNC or mapped drives to a centralized local folder. | Ensure formula integrity by recalculating all formulas after external link paths are updated.
// AI Prompts: Write a C# method that returns only ExternalLink objects whose DataSource is a UNC or mapped‑drive path. | Create a function that changes the DataSource of a list of ExternalLink objects to a specified folder and stores the original path in OriginalDataSource. | Generate code that logs a migration checklist for each external link, updates the link paths, calls workbook.CalculateFormula, and saves the workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook, scans Worksheets.ExternalLinks for UNC or mapped‑drive paths, prints a migration checklist for each link, updates the DataSource to a new local folder while preserving the original path, recalculates formulas, and saves the modified file.
class ExternalLinkMigrationChecklist
{
    static void Main()
    {
        // Path to the workbook that needs to be inspected
        string inputPath = "input.xlsx";

        // Path where the updated workbook will be saved
        string outputPath = "output.xlsx";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(inputPath);

        // Collect external links that point to network locations
        List<ExternalLink> networkLinks = new List<ExternalLink>();
        foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
        {
            string source = link.DataSource;
            if (IsNetworkPath(source))
            {
                networkLinks.Add(link);
            }
        }

        // Generate a migration checklist for each identified network link
        Console.WriteLine("=== Migration Checklist for Network External Links ===");
        foreach (ExternalLink link in networkLinks)
        {
            Console.WriteLine($"- External Link: {link.DataSource}");
            Console.WriteLine("  1. Verify that the network location is accessible.");
            Console.WriteLine("  2. If the workbook will be moved, copy the source file to a local/shared folder.");
            Console.WriteLine("  3. Update the DataSource to the new location (preserve original in OriginalDataSource).");
            Console.WriteLine("  4. Recalculate formulas to ensure values are refreshed.");
        }

        // Example: update all identified network links to a new local folder
        string newFolder = @"C:\MigratedExternalFiles\";
        foreach (ExternalLink link in networkLinks)
        {
            // Preserve the original path
            link.OriginalDataSource = link.DataSource;

            // Build the new path using the same file name
            string fileName = System.IO.Path.GetFileName(link.DataSource);
            string newPath = System.IO.Path.Combine(newFolder, fileName);

            // Update the link to point to the new location
            link.DataSource = newPath;
        }

        // Recalculate formulas after updating the external links
        workbook.CalculateFormula();

        // Save the modified workbook (save rule)
        workbook.Save(outputPath);
    }

    // Helper method to decide whether a path points to a network location
    static bool IsNetworkPath(string path)
    {
        // UNC paths start with double backslashes (e.g., \\server\share\file.xlsx)
        if (path.StartsWith(@"\\"))
            return true;

        // Simple heuristic for mapped drives (e.g., Z:\folder\file.xlsx)
        // Adjust this logic if you need stricter validation.
        if (path.Length > 2 && path[1] == ':' && (path[2] == '\\' || path[2] == '/'))
            return true;

        return false;
    }
}
