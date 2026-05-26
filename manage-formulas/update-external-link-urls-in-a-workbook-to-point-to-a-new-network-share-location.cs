using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UpdateExternalLinksDemo
    {
        // Adjust these paths as needed
        private const string InputFilePath = @"C:\Input\WorkbookWithLinks.xlsx";
        private const string OutputFilePath = @"C:\Output\WorkbookWithUpdatedLinks.xlsx";
        private const string OldNetworkShare = @"\\oldserver\share\";
        private const string NewNetworkShare = @"\\newserver\share\";

        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook processed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Verify input file exists
            if (!File.Exists(InputFilePath))
                throw new FileNotFoundException($"Input file not found: {InputFilePath}");

            // Load the workbook that contains external links
            Workbook workbook = new Workbook(InputFilePath);

            // Get the collection of external links
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Iterate through each external link and replace the old network share with the new one
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // Update the DataSource (the current link path)
                if (!string.IsNullOrEmpty(link.DataSource) &&
                    link.DataSource.StartsWith(OldNetworkShare, StringComparison.OrdinalIgnoreCase))
                {
                    string updatedPath = link.DataSource.Replace(OldNetworkShare, NewNetworkShare);
                    link.DataSource = updatedPath;
                }

                // Also update OriginalDataSource if it exists (preserves original stored path)
                if (!string.IsNullOrEmpty(link.OriginalDataSource) &&
                    link.OriginalDataSource.StartsWith(OldNetworkShare, StringComparison.OrdinalIgnoreCase))
                {
                    string updatedOriginal = link.OriginalDataSource.Replace(OldNetworkShare, NewNetworkShare);
                    link.OriginalDataSource = updatedOriginal;
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(OutputFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook with updated external link URLs
            workbook.Save(OutputFilePath);
        }
    }
}