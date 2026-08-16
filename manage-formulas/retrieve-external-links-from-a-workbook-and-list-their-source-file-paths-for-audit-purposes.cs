// Title: C# – List External Link File Paths in an Excel Workbook with Aspose.Cells
// Description: Loads an Excel workbook using Aspose.Cells, accesses the Worksheets.ExternalLinks collection, and enumerates each ExternalLink.DataSource to output the source file path, enabling quick audit of external references.
// Keywords: Aspose.Cells | C# | external links | DataSource | Excel audit | list link paths | Workbook.ExternalLinks | enumerate external references | file path extraction
// Common Searches: Aspose.Cells get external link paths | C# list external references in Excel | How to enumerate external links with Aspose.Cells | Retrieve DataSource of external links Aspose | Audit Excel external links C#
// Developer Intent: Extract every external workbook reference (file path) from an Excel file for auditing or reporting purposes.
// Use Cases: Create a compliance report that lists all workbooks referenced by the current Excel file. | Detect and flag broken or outdated external links before sharing the workbook. | Log external link paths to a database for version‑control and change‑tracking. | Generate a CSV inventory of external references for downstream processing.
// AI Prompts: Write a C# method using Aspose.Cells that returns a List<string> of all ExternalLink.DataSource values from a given Excel file. | Explain how to handle a workbook that contains no external links when performing an audit with Aspose.Cells. | Provide sample code that saves the extracted external link file paths to a CSV file for further analysis. | Show how to filter out duplicate external link paths while enumerating Workbook.ExternalLinks in C#.

using System;
using Aspose.Cells;

namespace ExternalLinksAudit
{
    // Loads an Excel workbook using Aspose.Cells, accesses the Worksheets.ExternalLinks collection, and enumerates each ExternalLink.DataSource to output the source file path, enabling quick audit of external references.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that needs to be audited
            string workbookPath = "InputWorkbook.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Access the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Check if there are any external links
            if (externalLinks.Count == 0)
            {
                Console.WriteLine("No external links found in the workbook.");
            }
            else
            {
                Console.WriteLine($"Found {externalLinks.Count} external link(s):");

                // Iterate through each external link and output its data source (file path)
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];
                    Console.WriteLine($"Link {i + 1}: {link.DataSource}");
                }
            }

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("AuditedWorkbook.xlsx"); // uses the provided save rule
        }
    }
}
