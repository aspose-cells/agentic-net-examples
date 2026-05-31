using System;
using Aspose.Cells;

namespace ExternalLinksAudit
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the collection of external links
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

            // Optionally, save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}