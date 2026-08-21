// Title: C# – Update Excel external link paths to a network share using Aspose.Cells
// Description: Loads a master workbook, scans its ExternalLinkCollection, replaces any data‑source paths that begin with an old folder prefix with a new network‑share prefix, updates both OriginalDataSource and DataSource, reloads the referenced workbooks, calls UpdateLinkedDataSource, recalculates formulas, and saves the workbook with corrected links.
// Keywords: Aspose.Cells external links | C# update Excel link path | network share workbook reference | ExternalLinkCollection replace folder prefix | .NET Excel linked data source | recalculate formulas Aspose.Cells
// Common Searches: change external workbook path Aspose.Cells C# | update Excel external links after moving files | replace folder prefix in Excel external references .NET | Aspose.Cells recalculate formulas after path change | load and refresh linked data sources C#
// Developer Intent: Modify a workbook’s external link formulas so they point to a new network‑share location and refresh the linked data.
// Use Cases: Migrate source files to a shared drive and automatically correct all external references in a master workbook. | Batch‑process multiple workbooks to replace an outdated folder prefix with a new UNC path. | Validate the presence of external workbooks after a path change, reload them, and recalculate formulas to maintain data integrity.
// AI Prompts: Write C# code with Aspose.Cells that scans a workbook’s ExternalLinkCollection, swaps an old folder prefix for a new network‑share prefix, and saves the updated file. | Create a method that updates external link data sources, loads each referenced workbook, calls UpdateLinkedDataSource, recalculates formulas, and handles missing files gracefully. | Explain the difference between OriginalDataSource and DataSource in Aspose.Cells when updating external links and how to ensure formulas are refreshed.

using System;
using System.IO;
using Aspose.Cells;

namespace UpdateExternalLinksDemo
{
    // Loads a master workbook, scans its ExternalLinkCollection, replaces any data‑source paths that begin with an old folder prefix with a new network‑share prefix, updates both OriginalDataSource and DataSource, reloads the referenced workbooks, calls UpdateLinkedDataSource, recalculates formulas, and saves the workbook with corrected links.
    class Program
    {
        static void Main()
        {
            // Paths and folder prefixes
            string mainWorkbookPath = @"C:\Temp\MainWorkbook.xlsx";
            string oldFolderPrefix = @"C:\OldData\";
            string newFolderPrefix = @"\\NetworkShare\NewData\";

            try
            {
                // Load the main workbook; create a new one if the file does not exist
                Workbook mainWorkbook;
                if (File.Exists(mainWorkbookPath))
                {
                    mainWorkbook = new Workbook(mainWorkbookPath);
                }
                else
                {
                    Console.WriteLine($"Main workbook not found at '{mainWorkbookPath}'. Creating a new workbook for demonstration.");
                    mainWorkbook = new Workbook(); // empty workbook
                }

                // Get external links collection
                ExternalLinkCollection externalLinks = mainWorkbook.Worksheets.ExternalLinks;

                // Update each external link's data source path
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    ExternalLink link = externalLinks[i];
                    string currentSource = !string.IsNullOrEmpty(link.OriginalDataSource)
                                            ? link.OriginalDataSource
                                            : link.DataSource;

                    if (currentSource.StartsWith(oldFolderPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        string updatedSource = newFolderPrefix + currentSource.Substring(oldFolderPrefix.Length);
                        link.OriginalDataSource = updatedSource;
                        link.DataSource = updatedSource;
                        Console.WriteLine($"Link {i} updated to: {updatedSource}");
                    }
                    else
                    {
                        Console.WriteLine($"Link {i} does not need updating: {currentSource}");
                    }
                }

                // Load external workbooks based on updated data sources
                Workbook[] externalWorkbooks = new Workbook[externalLinks.Count];
                for (int i = 0; i < externalLinks.Count; i++)
                {
                    string externalPath = externalLinks[i].DataSource;
                    if (File.Exists(externalPath))
                    {
                        externalWorkbooks[i] = new Workbook(externalPath);
                    }
                    else
                    {
                        Console.WriteLine($"Warning: External workbook not found at {externalPath}");
                        externalWorkbooks[i] = null;
                    }
                }

                // Remove null entries
                externalWorkbooks = Array.FindAll(externalWorkbooks, wb => wb != null);

                // Update linked data sources if any external workbooks were loaded
                if (externalWorkbooks.Length > 0)
                {
                    mainWorkbook.UpdateLinkedDataSource(externalWorkbooks);
                }

                // Recalculate formulas to reflect any changes
                mainWorkbook.CalculateFormula();

                // Save the updated workbook
                string outputPath = Path.Combine(Path.GetDirectoryName(mainWorkbookPath) ?? Environment.CurrentDirectory,
                                                "MainWorkbook_Updated.xlsx");
                mainWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with updated external links at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
