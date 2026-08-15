// Title: C# Example: Audit and Update External Link Paths in an Excel Workbook with Aspose.Cells
// Description: This C# snippet loads an Excel workbook, enumerates all external links, logs each link's original data source, replaces a specified old base URL with a new path, writes the updated value back, logs the change for audit purposes, and saves the workbook. Ideal for migration, compliance checks, and automated link management.
// Keywords: Aspose.Cells | C# | .NET | Excel external links | OriginalDataSource | audit external links | URL replacement | log link paths | workbook migration | compliance review | GitHub sample | code example
// Common Searches: how to read external links with Aspose.Cells C# | update external link paths in Excel programmatically | audit external data sources in a workbook using Aspose.Cells | replace old server URL in Excel external links C# | log original and new external link paths Aspose.Cells
// Developer Intent: Retrieve, modify, and record original and new external link paths in an Excel file.
// Use Cases: Create an audit trail before distributing a workbook to track link changes. | Migrate all external references from a legacy server to a new shared directory across multiple files. | Validate link updates during a regulatory compliance review. | Generate a report of external link sources for data governance. | Automate bulk correction of broken external links in a document repository.
// AI Prompts: Write a C# method using Aspose.Cells that extracts all external link URLs from a workbook and exports them to a CSV file. | Provide a reusable function that accepts oldBaseUrl and newBaseUrl strings and updates the OriginalDataSource of every external link in a given workbook. | Explain how to handle null or empty OriginalDataSource values safely while logging audit information with Aspose.Cells. | Generate a PowerShell script that calls a compiled C# program to batch‑process a folder of Excel files, updating external link paths as described. | Suggest unit tests for verifying that external link paths are correctly replaced and logged.

using System;
using Aspose.Cells;

// This C# snippet loads an Excel workbook, enumerates all external links, logs each link's original data source, replaces a specified old base URL with a new path, writes the updated value back, logs the change for audit purposes, and saves the workbook. Ideal for migration, compliance checks, and automated link management.
class ExternalLinkAudit
{
    static void Main()
    {
        // Load the workbook that contains external links
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each external link, log original and updated paths
        for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
        {
            ExternalLink link = workbook.Worksheets.ExternalLinks[i];

            // Capture the original stored data source
            string originalPath = link.OriginalDataSource;
            Console.WriteLine($"External Link {i} - Original Path: {originalPath}");

            // Example modification: replace an old base URL with a new one
            string updatedPath = originalPath.Replace(
                @"https://oldserver.com/",
                @"/shared/files/");

            // Apply the updated path back to the external link
            link.OriginalDataSource = updatedPath;

            // Log the updated path for audit
            Console.WriteLine($"External Link {i} - Updated Path: {updatedPath}");
        }

        // Save the workbook after modifications
        workbook.Save("output.xlsx");
    }
}
