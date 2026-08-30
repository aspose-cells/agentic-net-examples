// Title: How to programmatically change the folder path of external link data sources in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates over Workbook.Worksheets.ExternalLinks, and replaces any OriginalDataSource or DataSource that begins with a given old directory with a new directory, then saves the updated file. | Provide a C# example that checks the input file existence, updates external link paths, creates the output directory if needed, and includes proper exception handling while using Aspose.Cells.
// Common Searches: C# Aspose.Cells update external link file path in existing workbook | replace old folder path with new path for Excel external data source using Aspose.Cells .NET | how to modify ExternalLinkCollection OriginalDataSource programmatically in Aspose.Cells | Aspose.Cells change external link source directory for multiple links | save workbook after updating external link paths Aspose.Cells C#
// Tags: external link path update Aspose.Cells | Workbook.ExternalLinks modify data source C# | replace base directory in Excel external links .NET | Aspose.Cells external data source path change | save workbook after external link modification

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads a workbook, iterates through its ExternalLinkCollection, replaces any OriginalDataSource or DataSource that starts with a specified old base folder with a new folder path, ensures the output directory exists, and saves the modified workbook to a new file while handling errors.
    public class UpdateExternalLinkPathDemo
    {
        public static void Run()
        {
            try
            {
                // Path to the input workbook
                string inputPath = "input.xlsx";

                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains external links
                Workbook workbook = new Workbook(inputPath);

                // Define the part of the path to replace and the new path
                string oldBasePath = @"C:\OldFolder\";
                string newBasePath = @"D:\NewFolder\";

                // Iterate through all external links and update their stored data source paths
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    // Update OriginalDataSource if it matches the old base path
                    string original = externalLinks[i].OriginalDataSource;
                    if (!string.IsNullOrEmpty(original) && original.StartsWith(oldBasePath, StringComparison.OrdinalIgnoreCase))
                    {
                        externalLinks[i].OriginalDataSource = original.Replace(oldBasePath, newBasePath);
                    }

                    // Update DataSource if it differs and matches the old base path
                    string dataSource = externalLinks[i].DataSource;
                    if (!string.IsNullOrEmpty(dataSource) && dataSource.StartsWith(oldBasePath, StringComparison.OrdinalIgnoreCase))
                    {
                        externalLinks[i].DataSource = dataSource.Replace(oldBasePath, newBasePath);
                    }
                }

                // Path to the output workbook
                string outputPath = "output.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the modified external link paths
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            UpdateExternalLinkPathDemo.Run();
        }
    }
}
