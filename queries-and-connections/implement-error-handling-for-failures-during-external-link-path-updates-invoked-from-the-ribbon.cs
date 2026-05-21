using System;
using Aspose.Cells;

namespace ExternalLinkPathUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output workbook file paths
            string inputFile = "InputWorkbook.xlsx";
            string outputFile = "UpdatedWorkbook.xlsx";

            // Old and new base paths to replace in external link data sources
            string oldBasePath = @"C:\OldExternalLinks\";
            string newBasePath = @"D:\NewExternalLinks\";

            try
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputFile);

                // Get the collection of external links
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

                // Iterate through each external link and attempt to update its DataSource
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    try
                    {
                        ExternalLink link = externalLinks[i];

                        // Preserve the original data source for logging
                        string originalDataSource = link.DataSource;

                        // Perform the path replacement
                        string updatedDataSource = originalDataSource.Replace(oldBasePath, newBasePath);

                        // If no change occurred, skip assignment
                        if (!originalDataSource.Equals(updatedDataSource, StringComparison.OrdinalIgnoreCase))
                        {
                            link.DataSource = updatedDataSource;
                            Console.WriteLine($"Link {i} updated: '{originalDataSource}' -> '{updatedDataSource}'");
                        }
                        else
                        {
                            Console.WriteLine($"Link {i} unchanged (no matching base path).");
                        }
                    }
                    catch (Exception linkEx)
                    {
                        // Handle failures for a specific link without aborting the whole process
                        Console.WriteLine($"Error updating external link at index {i}: {linkEx.Message}");
                    }
                }

                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                // General error handling for load/save operations
                Console.WriteLine($"Failed to process workbook: {ex.Message}");
            }
        }
    }
}