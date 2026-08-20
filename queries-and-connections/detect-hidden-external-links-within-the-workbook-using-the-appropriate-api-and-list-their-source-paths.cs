// Title: Detect Hidden External Links in an Excel Workbook with Aspose.Cells for .NET
// Description: Loads an Excel file using Aspose.Cells, iterates the workbook's ExternalLinkCollection, and prints the DataSource of links where IsVisible is false. Also outputs a complete list of all external links with their visibility status.
// Keywords: Aspose.Cells | C# external links | Hidden external connections | ExternalLink.IsVisible | Excel workbook audit | DataSource enumeration | .NET Excel API
// Common Searches: find hidden external links Aspose.Cells C# | list invisible external connections in Excel workbook | ExternalLinkCollection IsVisible false example | retrieve external link data source Aspose.Cells | audit Excel file for concealed links .NET
// Developer Intent: Identify every external link in a workbook that is marked as hidden (IsVisible = false) and retrieve its source path.
// Use Cases: Perform a compliance check to ensure no concealed data connections are shipped with a spreadsheet. | Create a report of all hidden external references before publishing or sharing a workbook. | Validate the cleanup of a workbook by confirming that no hidden links remain.
// AI Prompts: Generate C# code with Aspose.Cells that removes all hidden external links after listing their DataSource values. | Write a method that returns a List<string> of DataSource strings for every invisible external link in a given workbook. | Explain how to change the IsVisible property to true for hidden external links using Aspose.Cells.

using System;
using Aspose.Cells;

namespace DetectHiddenExternalLinks
{
    // Loads an Excel file using Aspose.Cells, iterates the workbook's ExternalLinkCollection, and prints the DataSource of links where IsVisible is false. Also outputs a complete list of all external links with their visibility status.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the collection of external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Check if there are any external links at all
            if (externalLinks.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any external links.");
                return;
            }

            Console.WriteLine("Hidden external links (IsVisible == false):");

            // Iterate through all external links and list those that are hidden
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // IsVisible indicates whether the link is visible in Excel.
                // Hidden links have IsVisible == false.
                if (!link.IsVisible)
                {
                    Console.WriteLine($"- Link {i}: {link.DataSource}");
                }
            }

            // Optionally, you can also list all external links for reference
            Console.WriteLine("\nAll external links in the workbook:");
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];
                Console.WriteLine($"- Link {i}: {link.DataSource} (Visible: {link.IsVisible})");
            }

            // No changes are made, but if you need to save the workbook after processing:
            // workbook.Save("output.xlsx");
        }
    }
}
