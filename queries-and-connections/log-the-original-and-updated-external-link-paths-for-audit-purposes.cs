// Title: How to Audit and Rewrite External Link Paths in an Excel Workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to enumerate all ExternalLink objects in a workbook, print each link's OriginalDataSource, replace a specified URL prefix, and save the modified file. | Show a .NET snippet that updates the OriginalDataSource property of external links in an Excel file, logs the before‑and‑after paths to the console, and persists the changes with Aspose.Cells. | Demonstrate how to perform an audit of external data source paths in an Excel workbook, apply a custom path transformation, and write the results using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# iterate external links and change their data source path | How to replace part of an external link URL in an Excel file using Aspose.Cells .NET | Log original and updated external link paths while saving workbook with Aspose.Cells | Audit external data source paths in Excel workbook programmatically with Aspose.Cells | Update ExternalLink.OriginalDataSource for all worksheets in C#
// Tags: external link path replacement Aspose.Cells | modify OriginalDataSource property C# | audit external links Excel Aspose.Cells | enumerate ExternalLinkCollection .NET | save workbook after external link update Aspose.Cells

using System;
using Aspose.Cells;

namespace ExternalLinkAuditDemo
{
    // The sample loads 'input.xlsx', iterates through its ExternalLinkCollection, records each link's OriginalDataSource, replaces the 'https://oldserver.com/files/' segment with '/shared/files/', writes both original and updated paths to the console, assigns the new path back to the link, and saves the workbook as 'output.xlsx'.
    public class Program
    {
        public static void Main()
        {
            // Load the workbook that contains external links
            Workbook workbook = new Workbook("input.xlsx");

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link, log original and updated paths
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // Store the original data source path
                string originalPath = link.OriginalDataSource;

                // Define how the path should be transformed (example replacement)
                string updatedPath = originalPath.Replace(
                    @"https://oldserver.com/files/",
                    @"/shared/files/");

                // Apply the updated path back to the external link
                link.OriginalDataSource = updatedPath;

                // Log the audit information
                Console.WriteLine($"External Link {i}:");
                Console.WriteLine($"  Original Path: {originalPath}");
                Console.WriteLine($"  Updated Path : {updatedPath}");
            }

            // Save the workbook after modifications
            workbook.Save("output.xlsx");
        }
    }
}
