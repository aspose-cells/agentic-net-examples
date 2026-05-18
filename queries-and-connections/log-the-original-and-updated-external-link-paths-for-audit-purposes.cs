using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExternalLinkAuditDemo
    {
        public static void Run()
        {
            // Load the workbook that contains external links
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link, log original and updated paths
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // Capture the original data source path
                string originalPath = link.DataSource;

                // Example modification: replace a specific folder segment
                // (adjust the strings as needed for your scenario)
                string updatedPath = originalPath.Replace(
                    @"https://oldserver.com/files/",
                    @"/shared/files/");

                // Apply the updated path back to the external link
                link.DataSource = updatedPath;

                // Log both paths for audit purposes
                Console.WriteLine($"Link {i}:");
                Console.WriteLine($"  Original Path: {originalPath}");
                Console.WriteLine($"  Updated Path : {updatedPath}");
            }

            // Save the workbook with the updated external link paths
            workbook.Save("OutputWorkbook.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExternalLinkAuditDemo.Run();
        }
    }
}