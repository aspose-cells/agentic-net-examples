// Title: Change external link paths in an XLS workbook to a UNC network share using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xls file, iterates over its ExternalLinkCollection, replaces a given local folder segment in each link's OriginalDataSource and DataSource with a UNC network share path, and saves the workbook. | Show how to update both OriginalDataSource and DataSource properties of ExternalLink objects to point to a new network share while keeping the rest of the link configuration intact with Aspose.Cells.
// Common Searches: c# update external link UNC path in existing xls workbook using Aspose.Cells | how to replace local folder segment in external link paths of an Excel file programmatically | asp.net modify OriginalDataSource of external links in an Excel workbook | iterate ExternalLinkCollection and change DataSource to network share in .NET | change external data source path to shared folder in Aspose.Cells workbook
// Tags: externallink collection path replacement Aspose.Cells | update OriginalDataSource UNC Aspose.Cells | modify external link DataSource .NET | xls workbook external link path update | Aspose.Cells network share link adjustment

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkPathUpdate
{
    // The example loads an XLS workbook, walks through its ExternalLinkCollection, replaces a specified local folder segment in each link's OriginalDataSource and DataSource with a UNC network share path, and saves the modified workbook to a new file.
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            string inputPath = Path.GetFullPath("input.xls");
            Workbook workbook = new Workbook(inputPath);

            // Define the part of the original path to replace and the new network share path
            // Example: replace "C:\\Data\\Reports\\" with "\\\\Server\\Shared\\Reports\\"
            string oldPathSegment = @"C:\Data\Reports\";
            string newNetworkPath = @"\\Server\Shared\Reports\";

            // Iterate through all external links in the workbook and update their paths
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // Use OriginalDataSource to preserve the original stored path
                string originalSource = link.OriginalDataSource;

                // If the original source contains the old segment, replace it
                if (!string.IsNullOrEmpty(originalSource) && originalSource.Contains(oldPathSegment))
                {
                    string updatedSource = originalSource.Replace(oldPathSegment, newNetworkPath);
                    link.OriginalDataSource = updatedSource;

                    // Also update DataSource to keep the link functional
                    link.DataSource = updatedSource;
                }
            }

            // Save the modified workbook (replace with your desired output file path)
            string outputPath = Path.GetFullPath("output.xls");
            workbook.Save(outputPath);

            Console.WriteLine("External link paths have been updated and workbook saved to:");
            Console.WriteLine(outputPath);
        }
    }
}
