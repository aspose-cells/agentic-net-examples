using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExternalLinkAuditDemo
    {
        public static void Run()
        {
            const string inputPath = "InputWithExternalLinks.xlsx";
            const string outputPath = "OutputWithUpdatedLinks.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook that contains external links
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each external link in the workbook
                for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
                {
                    ExternalLink link = workbook.Worksheets.ExternalLinks[i];

                    // Capture the original external link path for auditing
                    string originalPath = link.OriginalDataSource;

                    // Example transformation: replace an old base URL with a new one
                    string updatedPath = originalPath.Replace(
                        @"https://oldserver.com/files/",
                        @"/shared/files/");

                    // Update the external link with the new path
                    link.OriginalDataSource = updatedPath;

                    // Log the original and updated paths
                    Console.WriteLine($"External Link {i}:");
                    Console.WriteLine($"  Original Path: {originalPath}");
                    Console.WriteLine($"  Updated Path : {updatedPath}");
                }

                // Save the workbook after updating the external links
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}