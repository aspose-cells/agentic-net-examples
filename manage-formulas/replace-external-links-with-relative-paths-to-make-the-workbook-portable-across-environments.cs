using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkRelativePath
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source workbook (absolute path)
                string sourceWorkbookPath = @"C:\Data\Reports\MainReport.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(sourceWorkbookPath))
                    throw new FileNotFoundException($"Source workbook not found: {sourceWorkbookPath}");

                // Load the workbook
                Workbook workbook = new Workbook(sourceWorkbookPath);

                // Directory of the workbook – this will be the base for relative paths
                string workbookDirectory = Path.GetDirectoryName(sourceWorkbookPath);

                // Iterate through all external links in the workbook
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];

                    // Original absolute path of the external link
                    string absolutePath = link.DataSource;

                    // If the path is absolute, convert it to a relative path
                    if (!string.IsNullOrEmpty(absolutePath) && Path.IsPathRooted(absolutePath))
                    {
                        string relativePath = Path.GetRelativePath(workbookDirectory, absolutePath);
                        link.DataSource = relativePath;          // Update the path used by formulas
                        link.OriginalDataSource = relativePath; // Keep original reference consistent
                    }
                }

                // Prepare output path
                string outputPath = @"C:\Data\Reports\MainReport_Portable.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Save the modified workbook
                workbook.Save(outputPath);

                Console.WriteLine("External links have been converted to relative paths.");
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}