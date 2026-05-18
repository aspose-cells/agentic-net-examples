using System;
using Aspose.Cells;

namespace AsposeCellsExternalLinkRemoval
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that may contain hidden external links
            Workbook workbook = new Workbook("input.xlsx");

            // Get the collection of external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Remove hidden external links (IsVisible == false)
            for (int i = externalLinks.Count - 1; i >= 0; i--)
            {
                ExternalLink link = externalLinks[i];
                if (!link.IsVisible) // hidden link
                {
                    externalLinks.RemoveAt(i);
                }
            }

            // Verify that no hidden external links remain
            bool hiddenLinkExists = false;
            foreach (ExternalLink link in externalLinks)
            {
                if (!link.IsVisible)
                {
                    hiddenLinkExists = true;
                    break;
                }
            }

            Console.WriteLine("Hidden external link present after removal: " + hiddenLinkExists);
            Console.WriteLine("Remaining external links count: " + externalLinks.Count);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}