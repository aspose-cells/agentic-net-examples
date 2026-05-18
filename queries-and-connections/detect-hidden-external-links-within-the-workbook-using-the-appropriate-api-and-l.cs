using System;
using Aspose.Cells;

namespace AsposeCellsExternalLinkDetection
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Check if the workbook contains any external links
            if (workbook.Worksheets.ExternalLinks.Count == 0)
            {
                Console.WriteLine("No external links found in the workbook.");
                return;
            }

            Console.WriteLine("External links detected:");

            // Iterate through all external links
            for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
            {
                ExternalLink link = workbook.Worksheets.ExternalLinks[i];

                // Determine if the link is hidden (not visible in Excel)
                bool isHidden = !link.IsVisible;

                // Output the link information
                Console.WriteLine($"Link #{i + 1}:");
                Console.WriteLine($"  Data Source : {link.DataSource}");
                Console.WriteLine($"  Hidden      : {isHidden}");
            }

            // Optionally, save the workbook after processing (if any modifications were made)
            // workbook.Save("output.xlsx");
        }
    }
}