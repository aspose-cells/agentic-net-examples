// Title: Update External Link URLs to a New UNC Path in an Excel Workbook with Aspose.Cells for .NET
// Description: Loads a workbook, scans its ExternalLinkCollection, replaces the old network‑share prefix in each link's DataSource and OriginalDataSource with a new UNC prefix, and saves the modified file.
// Keywords: Aspose.Cells external links | C# update Excel UNC path | change network share prefix | modify DataSource Aspose | OriginalDataSource update | .NET Excel external link batch | workbook link migration | replace external link URL | Excel workbook server move
// Common Searches: Aspose.Cells replace UNC path in external links | C# change external link source in Excel file | update network share prefix for workbook links | batch edit external links with Aspose.Cells | how to modify DataSource of external links .NET
// Developer Intent: Swap the old UNC share prefix for a new one across all external links in a workbook and persist the changes.
// Use Cases: Migrate Excel workbooks after moving a shared data server. | Automate link updates during a large‑scale file‑system reorganization. | Validate and enforce consistent external data sources before publishing workbooks.
// AI Prompts: Write C# code using Aspose.Cells that replaces a given old UNC prefix with a new one in every ExternalLink of an Excel workbook. | Create a reusable method that accepts input file path, old prefix, new prefix, and output path, then updates all external link URLs. | Explain how to programmatically confirm that DataSource and OriginalDataSource values were correctly rewritten after saving the workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, scans its ExternalLinkCollection, replaces the old network‑share prefix in each link's DataSource and OriginalDataSource with a new UNC prefix, and saves the modified file.
    class UpdateExternalLinks
    {
        static void Main()
        {
            try
            {
                // Input and output file paths
                string inputPath = @"C:\Data\MyWorkbook.xlsx";
                string outputPath = @"C:\Data\MyWorkbook_Updated.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains external links
                Workbook workbook = new Workbook(inputPath);

                // Define the old network share prefix and the new one
                string oldPrefix = @"\\oldserver\share\";
                string newPrefix = @"\\newserver\share\";

                // Get the collection of external links
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

                // Iterate through each external link and replace the old prefix with the new one
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];

                    // Update DataSource if it starts with the old prefix
                    if (!string.IsNullOrEmpty(link.DataSource) &&
                        link.DataSource.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        string updatedPath = newPrefix + link.DataSource.Substring(oldPrefix.Length);
                        link.DataSource = updatedPath;
                    }

                    // Update OriginalDataSource similarly (optional but ensures all stored paths are changed)
                    if (!string.IsNullOrEmpty(link.OriginalDataSource) &&
                        link.OriginalDataSource.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        string updatedOriginal = newPrefix + link.OriginalDataSource.Substring(oldPrefix.Length);
                        link.OriginalDataSource = updatedOriginal;
                    }
                }

                // Save the workbook with updated external link URLs
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
