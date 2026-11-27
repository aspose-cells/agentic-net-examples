// Title: C# program to update Excel external link paths with Aspose.Cells and gracefully handle missing or inaccessible files
// Common Searches: Aspose.Cells C# change external link datasource to new folder and skip missing files | How to verify external workbook file accessibility before updating DataSource with Aspose.Cells | C# update linked Excel file paths and handle UnauthorizedAccessException using Aspose.Cells | Refresh external data sources after moving linked Excel files in Aspose.Cells .NET | Backup Excel workbook while updating external links with Aspose.Cells
// Tags: modify external link path Aspose.Cells | check external file existence C# | handle unauthorized access exception Aspose.Cells | refresh linked data sources Aspose.Cells | save workbook with backup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExternalLinkPathUpdater
{
    // The example loads a workbook, iterates through its ExternalLinks collection, extracts each link's file name, and attempts to replace the DataSource with a new location. It validates that the target file exists and can be opened for reading, handling missing files, access‑denied errors, and other I/O issues. Valid links are refreshed using UpdateLinkedDataSource, and the workbook is saved under a new name to preserve the original file.
    class Program
    {
        /// <param name="args">
        /// args[0] - Path to the workbook that contains external links.
        /// args[1] - New directory where the external files are expected to reside.
        /// </param>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ExternalLinkPathUpdater <WorkbookPath> <NewExternalFilesDirectory>");
                return;
            }

            string workbookPath = args[0];
            string newExternalDir = args[1];

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            if (!Directory.Exists(newExternalDir))
            {
                Console.WriteLine($"Error: Target directory for external files does not exist: '{newExternalDir}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);
                Console.WriteLine($"Loaded workbook: {Path.GetFileName(workbookPath)}");

                // Process each external link
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
                if (externalLinks.Count == 0)
                {
                    Console.WriteLine("No external links found in the workbook.");
                }
                else
                {
                    for (int i = 0; i < externalLinks.Count; i++)
                    {
                        ExternalLink link = externalLinks[i];
                        string oldDataSource = link.DataSource;
                        Console.WriteLine($"Processing external link #{i + 1}: {oldDataSource}");

                        // Determine the file name from the old data source (could be a full path or just a file name)
                        string fileName = Path.GetFileName(oldDataSource);
                        if (string.IsNullOrEmpty(fileName))
                        {
                            Console.WriteLine($"  Warning: Unable to extract file name from data source '{oldDataSource}'. Skipping.");
                            continue;
                        }

                        string proposedPath = Path.Combine(newExternalDir, fileName);

                        // Verify that the target file exists and is accessible
                        if (!File.Exists(proposedPath))
                        {
                            Console.WriteLine($"  Error: Target external file does not exist: '{proposedPath}'. Link will remain unchanged.");
                            continue;
                        }

                        try
                        {
                            // Attempt to open the file with read permission to ensure accessibility
                            using (FileStream fs = File.Open(proposedPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                            {
                                // If we reach here, the file is accessible
                            }

                            // Update the DataSource property
                            link.DataSource = proposedPath;
                            Console.WriteLine($"  Updated link to new path: {proposedPath}");
                        }
                        catch (UnauthorizedAccessException uae)
                        {
                            Console.WriteLine($"  Error: Access denied to file '{proposedPath}'. {uae.Message}");
                        }
                        catch (IOException ioex)
                        {
                            Console.WriteLine($"  Error: I/O problem with file '{proposedPath}'. {ioex.Message}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Unexpected error while updating link: {ex.Message}");
                        }
                    }

                    // After updating paths, attempt to refresh data from the external sources
                    try
                    {
                        // Collect the external workbooks that correspond to the updated links
                        Workbook[] externalWorkbooks = new Workbook[externalLinks.Count];
                        for (int i = 0; i < externalLinks.Count; i++)
                        {
                            string ds = externalLinks[i].DataSource;
                            if (File.Exists(ds))
                            {
                                externalWorkbooks[i] = new Workbook(ds);
                            }
                            else
                            {
                                externalWorkbooks[i] = null; // Will be ignored by UpdateLinkedDataSource
                            }
                        }

                        // Remove null entries before calling the method
                        externalWorkbooks = Array.FindAll(externalWorkbooks, w => w != null);
                        if (externalWorkbooks.Length > 0)
                        {
                            workbook.UpdateLinkedDataSource(externalWorkbooks);
                            Console.WriteLine("External data sources refreshed successfully.");
                        }
                        else
                        {
                            Console.WriteLine("No valid external workbooks found to refresh data.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error while refreshing external data sources: {ex.Message}");
                    }
                }

                // Save the workbook with a new name to avoid overwriting the original
                string backupPath = Path.Combine(Path.GetDirectoryName(workbookPath),
                    Path.GetFileNameWithoutExtension(workbookPath) + "_Updated" + Path.GetExtension(workbookPath));

                workbook.Save(backupPath);
                Console.WriteLine($"Workbook saved with updated links to: {backupPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
