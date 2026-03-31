using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHiddenExternalLinks
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be inspected
            string workbookPath = "input.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // List to hold information about hidden external links
            List<string> hiddenLinksInfo = new List<string>();

            // Iterate through all external links and check visibility
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // IsVisible indicates whether the link is visible in Excel.
                // Hidden links have IsVisible == false.
                if (!link.IsVisible)
                {
                    string info = $"Index: {i}, DataSource: {link.DataSource}, IsReferred: {link.IsReferred}";
                    hiddenLinksInfo.Add(info);
                    Console.WriteLine(info);
                }
            }

            // Output result summary
            if (hiddenLinksInfo.Count == 0)
            {
                Console.WriteLine("No hidden external links found.");
            }
            else
            {
                Console.WriteLine($"{hiddenLinksInfo.Count} hidden external link(s) detected.");
            }

            // Write findings to a text file (output via console)
            string outputFile = "HiddenExternalLinks.txt";
            File.WriteAllLines(outputFile, hiddenLinksInfo);
            Console.WriteLine($"Findings written to: {outputFile}");

            // Save the workbook if any modifications were made (save rule)
            // In this scenario we only read, but the save call demonstrates the rule usage.
            workbook.Save("output.xlsx");
        }
    }
}