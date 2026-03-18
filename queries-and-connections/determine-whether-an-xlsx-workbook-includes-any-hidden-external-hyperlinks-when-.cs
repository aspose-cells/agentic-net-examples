using System;
using Aspose.Cells;

class CheckHiddenExternalLinks
{
    static void Main()
    {
        // Load the workbook with default settings
        Workbook workbook = new Workbook("input.xlsx");

        bool hasHiddenExternalLinks = false;

        // Iterate through all external links in the workbook
        foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
        {
            // External links that are not visible in Excel are considered hidden
            if (!link.IsVisible)
            {
                hasHiddenExternalLinks = true;
                Console.WriteLine($"Hidden external link detected: {link.DataSource}");
            }
        }

        Console.WriteLine($"Workbook contains hidden external links: {hasHiddenExternalLinks}");
    }
}