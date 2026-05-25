using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchExternalLinkUpdater
{
    class Program
    {
        static void Main()
        {
            // Define the source workbook files to process
            string[] sourceFiles = new string[]
            {
                @"C:\Data\Report1.xlsx",
                @"C:\Data\Report2.xlsx",
                @"C:\Data\Report3.xlsx"
            };

            // Define the folder where updated workbooks will be saved
            string outputFolder = @"C:\Data\Updated";

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // Mapping function to convert old external link paths to new ones
            Func<string, string> mapOldPathToNew = oldPath =>
            {
                string oldBase = @"\\oldserver\shared\";
                string newBase = @"\\newserver\shared\";
                return oldPath.Replace(oldBase, newBase);
            };

            // Process each workbook
            foreach (string sourcePath in sourceFiles)
            {
                // Skip missing files and continue with the next one
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (lifecycle: load)
                    using (Workbook wb = new Workbook(sourcePath))
                    {
                        // Get the collection of external links
                        ExternalLinkCollection externalLinks = wb.Worksheets.ExternalLinks;

                        // Update each external link's data source using the mapping function
                        for (int i = 0; i < externalLinks.Count; i++)
                        {
                            ExternalLink link = externalLinks[i];

                            // OriginalDataSource holds the stored path; update it
                            string original = link.OriginalDataSource;
                            string updated = mapOldPathToNew(original);
                            link.OriginalDataSource = updated;

                            // Also update the DataSource property if needed (it reflects the current link)
                            link.DataSource = updated;
                        }

                        // Optionally recalculate formulas after updating links
                        wb.CalculateFormula();

                        // Build the output file path (same name, different folder)
                        string fileName = Path.GetFileName(sourcePath);
                        string outputPath = Path.Combine(outputFolder, fileName);

                        // Save the updated workbook (lifecycle: save)
                        wb.Save(outputPath);

                        Console.WriteLine($"Processed and saved: {outputPath}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{sourcePath}': {ex.Message}");
                }
            }
        }
    }
}