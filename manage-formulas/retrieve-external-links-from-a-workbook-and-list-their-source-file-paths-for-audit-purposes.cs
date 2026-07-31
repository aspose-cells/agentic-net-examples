// Title: C# – List External Link Paths in an Excel Workbook Using Aspose.Cells
// Description: Loads an Excel file with Aspose.Cells, accesses the Worksheets.ExternalLinks collection, and prints each link's current DataSource and OriginalDataSource paths for audit and compliance checks.
// Keywords: Aspose.Cells external links | C# list DataSource paths | Excel workbook audit | ExternalLink OriginalDataSource | retrieve external references .NET
// Common Searches: enumerate external links Aspose.Cells C# | get source file path of Excel external reference | list workbook external links .NET | audit Excel external links programmatically
// Developer Intent: Extract every external reference from a workbook and display both its active and original file locations to verify link integrity.
// Use Cases: Detect broken or moved external references before publishing a spreadsheet. | Create a compliance report that lists all linked source files in a workbook. | Automate path validation or migration of external links during data migration projects.
// AI Prompts: Generate C# code that collects all ExternalLink.DataSource values into a List<string> using Aspose.Cells. | Write a method that returns a Dictionary<int, string> mapping each external link index to its OriginalDataSource. | Show how to replace every ExternalLink.DataSource with a new base directory while preserving OriginalDataSource values.

using System;
using Aspose.Cells;

namespace ExternalLinksAudit
{
    // Loads an Excel file with Aspose.Cells, accesses the Worksheets.ExternalLinks collection, and prints each link's current DataSource and OriginalDataSource paths for audit and compliance checks.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that needs to be audited
            string workbookPath = "input.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Get the collection of external links from the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Check if there are any external links
            if (externalLinks.Count == 0)
            {
                Console.WriteLine("No external links found in the workbook.");
            }
            else
            {
                Console.WriteLine($"Found {externalLinks.Count} external link(s):");

                // Iterate through each external link and output its source file path
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];

                    // DataSource holds the current file path of the external link
                    Console.WriteLine($"Link {i + 1}: DataSource = {link.DataSource}");

                    // OriginalDataSource holds the stored original path (if it was modified)
                    Console.WriteLine($"        OriginalDataSource = {link.OriginalDataSource}");
                }
            }

            // (Optional) Save the workbook after audit if any modifications were made
            // workbook.Save("output_audited.xlsx");
        }
    }
}
