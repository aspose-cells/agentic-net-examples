// Title: How to update external link file paths to a network share in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells update external link source file path after moving workbook to network share | replace local file references in Excel external links with UNC paths using Aspose.Cells .NET | modify ExternalLinkCollection SourceFileName property programmatically in C# | bulk remap external workbook links to network location with Aspose.Cells | handle missing external link files after migration using Aspose.Cells C#
// Tags: Aspose.Cells external link path update | C# ExternalLinkCollection SourceFileName modification | Excel external links UNC path migration | Aspose.Cells workbook external link remapping | dynamic property access Aspose.Cells external link

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, defines a dictionary that maps old local source file paths to new UNC network share locations, iterates through the workbook's ExternalLinkCollection, and updates each link's SourceFileName property (using dynamic access) when a matching mapping exists, then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = @"C:\Reports\Report.xlsx";
            const string outputPath = @"C:\Reports\Report_Updated.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Input workbook not found.", inputPath);

            // Load the workbook that contains external link formulas
            Workbook workbook = new Workbook(inputPath);

            // Mapping of old source file paths to new network share locations
            var pathMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { @"C:\Data\Source1.xlsx", @"\\NetworkShare\Data\Source1.xlsx" },
                { @"C:\Data\Source2.xlsx", @"\\NetworkShare\Data\Source2.xlsx" }
                // Add more mappings as needed
            };

            // Access the collection of external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link and update its file path if a mapping exists
            foreach (ExternalLink link in externalLinks)
            {
                try
                {
                    // Use dynamic to access SourceFileName (property may vary across versions)
                    dynamic dynLink = link;
                    string oldPath = dynLink.SourceFileName as string;

                    if (oldPath != null && pathMap.TryGetValue(oldPath, out string newPath))
                    {
                        dynLink.SourceFileName = newPath;
                    }
                }
                catch (Exception exLink)
                {
                    Console.WriteLine($"Failed to process an external link: {exLink.Message}");
                }
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the updated workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
