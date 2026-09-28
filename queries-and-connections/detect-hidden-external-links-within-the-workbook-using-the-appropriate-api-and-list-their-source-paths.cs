// Title: List hidden external links and their data source paths in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, iterates over Workbook.Worksheets.ExternalLinks, and prints the DataSource of each link where IsVisible is false. | Create a .NET console example that counts all external links in a workbook and displays each link’s visibility status (Visible/Hidden) together with its data source. | Write a C# program using Aspose.Cells that extracts and logs the source paths of hidden external links from an Excel file without modifying the workbook.
// Common Searches: how to retrieve hidden external link paths from an Excel file using Aspose.Cells C# | Aspose.Cells list external links and check IsVisible property in .NET | C# code to enumerate external links in a workbook and identify invisible links | detect and log hidden external links in .xlsx with Aspose.Cells for .NET | Aspose.Cells external link collection example for hidden links
// Tags: Aspose.Cells enumerate external links .xlsx | C# detect hidden external links Aspose.Cells | ExternalLinkCollection IsVisible check | list external link data source Aspose.Cells | Excel workbook hidden link extraction .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExternalLinkDetection
{
    // // Loads an Excel workbook, accesses its Worksheets.ExternalLinks collection, and prints each link’s DataSource along with its visibility status, enabling detection of hidden external links.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be inspected
            string workbookPath = "input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(workbookPath);

            // Access the collection of external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Check if any external links exist
            if (externalLinks.Count == 0)
            {
                Console.WriteLine("No external links found in the workbook.");
                return;
            }

            Console.WriteLine($"Total external links found: {externalLinks.Count}");
            Console.WriteLine("Listing external link source paths (including hidden links):");

            // Iterate through each external link and output its data source.
            // Hidden links are those where IsVisible is false.
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];
                string visibility = link.IsVisible ? "Visible" : "Hidden";
                Console.WriteLine($"Link {i + 1}:");
                Console.WriteLine($"  Data Source : {link.DataSource}");
                Console.WriteLine($"  Visibility  : {visibility}");
            }

            // No modifications are made, so no need to save (lifecycle rule: save not required here)
        }
    }
}
