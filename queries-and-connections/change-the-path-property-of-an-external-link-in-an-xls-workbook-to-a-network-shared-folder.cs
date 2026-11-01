// Title: How to change the SourceFullName of external workbook links to a UNC network share in an XLS file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xls workbook with Aspose.Cells, enumerates its Worksheets.ExternalLinks collection, and updates each link's SourceFullName to a given UNC path. | Demonstrate using .NET reflection to set the read‑only SourceFullName property of an Aspose.Cells ExternalLink object to a network shared folder location.
// Common Searches: C# Aspose.Cells change external link source to UNC path in .xls workbook | How to update external workbook references to a network share using Aspose.Cells | Set SourceFullName property for external links in legacy Excel file with Aspose.Cells | Programmatically modify external link paths in Excel XLS files via .NET | Using reflection to edit Aspose.Cells external link target folder
// Tags: Aspose.Cells set external link UNC path | C# modify external workbook reference in XLS | update external link source for legacy Excel file | reflection assign SourceFullName Aspose.Cells | network share path for Excel external link

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example loads an .xls workbook, accesses its ExternalLinkCollection, uses reflection to assign a new UNC network share path to each link's SourceFullName property, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xls";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // New network shared folder path for the external link
            string newNetworkPath = @"\\Server\SharedFolder\ExternalFile.xlsx";

            // Update the source file name of each external workbook link
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            foreach (ExternalLink link in externalLinks)
            {
                // Use reflection to set SourceFullName if the property exists
                PropertyInfo prop = link.GetType().GetProperty("SourceFullName", BindingFlags.Public | BindingFlags.Instance);
                if (prop != null && prop.CanWrite)
                {
                    try
                    {
                        prop.SetValue(link, newNetworkPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to set SourceFullName for a link: {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputPath = "output.xls";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
