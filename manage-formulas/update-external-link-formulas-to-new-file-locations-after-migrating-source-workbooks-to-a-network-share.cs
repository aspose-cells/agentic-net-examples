using System;
using System.IO;
using Aspose.Cells;

namespace UpdateExternalLinksDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the workbook that contains external links (old location)
                string workbookPath = @"C:\OldFolder\MainWorkbook.xlsx";

                // Verify the source workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Source workbook not found: {workbookPath}");
                    return;
                }

                // New network share folder where the source workbooks have been moved
                string newNetworkFolder = @"\\NetworkShare\NewFolder";

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Get the collection of external links in the workbook
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

                // Update each external link to point to the new location
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];

                    // Current data source (could be a full path or just a file name)
                    string currentSource = link.DataSource;

                    // Extract only the file name part (e.g., "SourceData.xlsx")
                    string fileName = Path.GetFileName(currentSource);

                    // Build the new full path on the network share
                    string updatedSource = Path.Combine(newNetworkFolder, fileName);

                    // Update the DataSource and OriginalDataSource properties
                    link.DataSource = updatedSource;
                    link.OriginalDataSource = updatedSource;
                }

                // Save the updated workbook
                string outputPath = @"C:\UpdatedFolder\MainWorkbook_Updated.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine("External links have been updated and workbook saved to:");
                Console.WriteLine(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}