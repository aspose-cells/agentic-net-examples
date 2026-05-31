using System;
using Aspose.Cells;

namespace AsposeCellsExternalLinkDetection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be inspected
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(inputPath);

            // Get the collection of external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // If there are no external links, inform the user and exit
            if (externalLinks.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any external links.");
                return;
            }

            Console.WriteLine($"Total external links found: {externalLinks.Count}");
            Console.WriteLine();

            // Iterate through each external link and display its details
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // DataSource holds the path to the linked file
                string dataSource = link.DataSource;

                // IsVisible indicates whether the link is visible in Excel (false may imply a hidden link)
                bool isVisible = link.IsVisible;

                // IsReferred indicates whether the link is actually referenced by any formula
                bool isReferred = link.IsReferred;

                Console.WriteLine($"Link #{i + 1}");
                Console.WriteLine($"  Data Source   : {dataSource}");
                Console.WriteLine($"  Visible       : {isVisible}");
                Console.WriteLine($"  Referenced    : {isReferred}");
                Console.WriteLine();
            }

            // Additionally, scan all cells to find formulas that contain hidden external links
            Console.WriteLine("Scanning cells for hidden external links (ContainsExternalLink = true)...");
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    // Only formula cells can contain external links
                    if (cell.IsFormula && cell.ContainsExternalLink)
                    {
                        // If the external link referenced by this cell is not visible, treat it as hidden
                        // (We rely on the ExternalLink collection to determine visibility)
                        foreach (ExternalLink link in externalLinks)
                        {
                            if (!link.IsVisible && cell.Formula.Contains(link.DataSource))
                            {
                                Console.WriteLine($"Hidden link found in sheet '{sheet.Name}', cell {cell.Name}:");
                                Console.WriteLine($"  Formula   : {cell.Formula}");
                                Console.WriteLine($"  DataSource: {link.DataSource}");
                            }
                        }
                    }
                }
            }

            // (Optional) Save the workbook after inspection – demonstrates the save rule
            // workbook.Save("InspectedWorkbook.xlsx");
        }
    }
}