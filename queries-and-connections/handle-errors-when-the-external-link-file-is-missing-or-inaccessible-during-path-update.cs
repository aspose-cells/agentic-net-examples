using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExternalLinkPathUpdateWithErrorHandling
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Get external links collection
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate backwards to safely remove items
            for (int i = externalLinks.Count - 1; i >= 0; i--)
            {
                ExternalLink link = externalLinks[i];
                string linkPath = link.DataSource;

                // Resolve relative paths using workbook's folder
                if (!Path.IsPathRooted(linkPath) && !string.IsNullOrEmpty(workbook.AbsolutePath))
                {
                    linkPath = Path.Combine(workbook.AbsolutePath, linkPath);
                }

                // Check if the external file exists
                if (!File.Exists(linkPath))
                {
                    Console.WriteLine($"Missing external file: {linkPath}");
                    // Remove link and update formulas to local references
                    externalLinks.RemoveAt(i, updateReferencesAsLocal: true);
                    Console.WriteLine($"Removed external link at index {i} and updated references.");
                }
                else
                {
                    // Example path update (e.g., C:\Data\ -> D:\Data\)
                    string updatedPath = linkPath.Replace(@"C:\Data\", @"D:\Data\");
                    if (!string.Equals(updatedPath, linkPath, StringComparison.OrdinalIgnoreCase))
                    {
                        link.DataSource = updatedPath;
                        Console.WriteLine($"Updated external link path to: {updatedPath}");
                    }
                }
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}