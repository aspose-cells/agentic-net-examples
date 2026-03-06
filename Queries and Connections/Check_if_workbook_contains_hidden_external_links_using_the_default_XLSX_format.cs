using System;
using Aspose.Cells;

namespace AsposeCellsHiddenExternalLinksDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook (XLSX format by default)
            string workbookPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Flag to indicate presence of hidden external links
            bool hasHiddenExternalLinks = false;

            // Iterate through each external link and check its visibility
            foreach (ExternalLink link in externalLinks)
            {
                // IsVisible == false means the link is hidden in Excel
                if (!link.IsVisible)
                {
                    hasHiddenExternalLinks = true;
                    break;
                }
            }

            // Output the result
            if (hasHiddenExternalLinks)
            {
                Console.WriteLine("The workbook contains hidden external links.");
            }
            else if (externalLinks.Count > 0)
            {
                Console.WriteLine("The workbook contains external links, but none are hidden.");
            }
            else
            {
                Console.WriteLine("The workbook does not contain any external links.");
            }

            // No modifications are made, so no need to save the workbook.
        }
    }
}