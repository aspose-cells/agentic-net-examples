using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ChangeExternalLinkPath
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains external links
                Workbook workbook = new Workbook(inputPath);

                // Define the part of the old path to be replaced and the new path segment
                string oldPathSegment = @"C:\OldFolder\";
                string newPathSegment = @"D:\NewFolder\";

                // Iterate through all external links in the workbook
                for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
                {
                    ExternalLink link = workbook.Worksheets.ExternalLinks[i];

                    // Use OriginalDataSource (stored data source) to modify the path
                    // If OriginalDataSource is empty, fall back to DataSource
                    string currentSource = !string.IsNullOrEmpty(link.OriginalDataSource)
                                            ? link.OriginalDataSource
                                            : link.DataSource;

                    if (string.IsNullOrEmpty(currentSource))
                        continue;

                    // Replace the old path segment with the new one
                    string updatedSource = currentSource.Replace(oldPathSegment, newPathSegment);

                    // Assign the modified path back to the appropriate property
                    if (!string.IsNullOrEmpty(link.OriginalDataSource))
                        link.OriginalDataSource = updatedSource;
                    else
                        link.DataSource = updatedSource;
                }

                // Save the workbook with the updated external link paths
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ChangeExternalLinkPath.Run();
        }
    }
}