// Title: How to list external link source file paths in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates through each worksheet, and prints the SourceFilePath of every external link. | Show how to use reflection to safely access the ExternalLinks collection across different Aspose.Cells versions and retrieve each link's source file path. | Create a console utility that counts external references in an Excel workbook and outputs each reference's source file location using Aspose.Cells for .NET. | Provide robust error handling for missing workbook files while auditing external links with Aspose.Cells.
// Common Searches: aspocells c# get external link source file path from workbook | list all external references in an Excel file using Aspose.Cells .NET | how to enumerate external links in Excel with Aspose.Cells and reflection | audit external link paths in .xlsx using Aspose.Cells C# console app | retrieve source file paths of external links in Excel via Aspose.Cells API
// Tags: Aspose.Cells external links enumeration | C# retrieve external link source path | Excel workbook audit external references .NET | reflection access ExternalLinks Aspose.Cells | count external links in .xlsx using Aspose.Cells

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

// // Loads an Excel workbook, uses reflection to access each worksheet's ExternalLinks collection, extracts the SourceFilePath of each external link, prints the paths with incremental numbering, and reports the total count of external links found.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook to be audited
            string workbookPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            int totalLinks = 0;
            int linkNumber = 1;

            // Iterate through worksheets and use reflection to access external links (compatible with multiple Aspose.Cells versions)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var externalLinksProp = sheet.GetType().GetProperty("ExternalLinks");
                if (externalLinksProp == null) continue; // Property not available in this version

                var externalLinks = externalLinksProp.GetValue(sheet) as IEnumerable;
                if (externalLinks == null) continue;

                foreach (var linkObj in externalLinks)
                {
                    totalLinks++;

                    // Retrieve the source file path via reflection
                    var sourcePathProp = linkObj.GetType().GetProperty("SourceFilePath");
                    string sourcePath = sourcePathProp?.GetValue(linkObj) as string ?? "N/A";

                    Console.WriteLine($"Link {linkNumber}: {sourcePath}");
                    linkNumber++;
                }
            }

            // Output the total number of external links found
            Console.WriteLine($"External links count: {totalLinks}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
