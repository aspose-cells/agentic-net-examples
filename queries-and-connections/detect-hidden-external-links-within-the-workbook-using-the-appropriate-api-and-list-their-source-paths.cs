using System;
using Aspose.Cells;

class DetectHiddenExternalLinks
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of external links in the workbook
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // If there are no external links, inform the user and exit
        if (externalLinks.Count == 0)
        {
            Console.WriteLine("No external links found in the workbook.");
            return;
        }

        // Iterate through the external links and list those that are hidden (IsVisible == false)
        bool hiddenFound = false;
        Console.WriteLine("Hidden external links (IsVisible = false):");
        for (int i = 0; i < externalLinks.Count; i++)
        {
            ExternalLink link = externalLinks[i];
            if (!link.IsVisible)
            {
                hiddenFound = true;
                Console.WriteLine($"Link {i}: {link.DataSource}");
            }
        }

        if (!hiddenFound)
        {
            Console.WriteLine("No hidden external links detected.");
        }
    }
}