// Title: Batch Update External Link Formulas in Multiple Excel Workbooks with Aspose.Cells (C#)
// Description: A C# console app that scans a folder of .xlsx files, loads each workbook with Aspose.Cells, replaces every external link’s DataSource using a user‑defined mapping, recalculates all formulas, and saves the updated files to a target directory while preserving original names.
// Keywords: Aspose.Cells external link update | batch Excel link replacement C# | replace external data source path .NET | recalculate formulas programmatically | process multiple workbooks Aspose | Excel automation external references | update workbook links in bulk | C# Excel file batch processing
// Common Searches: how to change external link paths in many Excel files using Aspose.Cells | batch update external references in .xlsx with C# | Aspose.Cells recalculate formulas after link change | C# code to replace external data source in Excel workbooks | process multiple workbooks external links Aspose
// Developer Intent: Replace old external link paths with new locations across a set of Excel workbooks and save the refreshed files.
// Use Cases: Migrate financial models to a new data warehouse by updating source file references in all linked workbooks. | Automate spreadsheet relocation after a server folder restructure, ensuring formulas point to the new files. | Refresh reporting dashboards after moving source data to a different drive, with automatic formula recalculation.
// AI Prompts: Write C# code that uses Aspose.Cells to map old external link filenames to new absolute paths, update the links, recalculate formulas, and save each workbook. | Show how to log any external links that do not match the provided mapping before saving the workbook. | Provide a verification routine that confirms every external link in a workbook has been updated to a new location.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchExternalLinkUpdater
{
    // A C# console app that scans a folder of .xlsx files, loads each workbook with Aspose.Cells, replaces every external link’s DataSource using a user‑defined mapping, recalculates all formulas, and saves the updated files to a target directory while preserving original names.
    class Program
    {
        static void Main()
        {
            // Folder containing the workbooks to process
            string inputFolder = @"C:\InputWorkbooks";
            // Folder where the updated workbooks will be saved
            string outputFolder = @"C:\UpdatedWorkbooks";

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // Mapping of old external link file names to new file locations
            // Key: old file name (as it appears in the external link)
            // Value: new absolute file path to replace with
            var linkMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "OldDataSource1.xlsx", @"D:\NewDataSources\NewDataSource1.xlsx" },
                { "OldDataSource2.xlsx", @"D:\NewDataSources\NewDataSource2.xlsx" }
                // Add more mappings as required
            };

            // Process each workbook file in the input folder
            foreach (string workbookPath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Load the workbook
                Workbook wb = new Workbook(workbookPath);

                // Update each external link in the workbook
                ExternalLinkCollection externalLinks = wb.Worksheets.ExternalLinks;
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];
                    string originalSource = link.DataSource; // Current external link path

                    // Determine if the link matches any entry in the mapping dictionary
                    foreach (var kvp in linkMapping)
                    {
                        if (originalSource.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace the old file name with the new absolute path
                            string updatedSource = originalSource.Replace(kvp.Key, kvp.Value, StringComparison.OrdinalIgnoreCase);
                            link.DataSource = updatedSource;
                            break; // Exit the inner loop once a match is found
                        }
                    }
                }

                // Recalculate formulas so that they reflect the new external data
                wb.CalculateFormula();

                // Save the updated workbook to the output folder, preserving the original file name
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(workbookPath));
                wb.Save(outputPath);
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
