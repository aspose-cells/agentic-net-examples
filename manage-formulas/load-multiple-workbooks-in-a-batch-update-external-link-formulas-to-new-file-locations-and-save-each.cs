// Title: Batch update external link formulas in multiple Excel workbooks with Aspose.Cells for .NET (C#)
// Description: A C# console example that loads a list of Excel files, maps old external link paths to new locations, updates each workbook's Worksheets.ExternalLinks DataSource, recalculates all formulas, and saves the changes. Ideal for automating link migration across many workbooks.
// Keywords: Aspose.Cells | C# | .NET | Excel external links | batch update | DataSource replacement | recalculate formulas | load workbook | save workbook | automation | Excel link migration
// Common Searches: Aspose.Cells batch update external links C# | Change external reference paths in multiple Excel files | Recalculate formulas after updating external links with Aspose | C# script to replace Excel external data source | Automate external link migration in .NET
// Developer Intent: Programmatically replace old external link paths with new ones in a set of Excel workbooks and refresh formulas using Aspose.Cells.
// Use Cases: Migrate linked data sources to a new server for all financial reporting workbooks. | Standardize external reference paths after reorganizing shared data folders. | Update chart or pivot table sources across departmental spreadsheets in one run. | Prepare legacy workbooks for a cloud‑based data repository. | Automate compliance checks by ensuring all external links point to approved files.
// AI Prompts: Write a C# snippet that reads a collection of Excel files and updates specified external link paths using Aspose.Cells. | Show how to iterate over Worksheets.ExternalLinks, change each DataSource, and recalculate formulas in Aspose.Cells. | Explain error handling for missing files when batch‑updating external links in multiple workbooks.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchExternalLinkUpdater
{
    // A C# console example that loads a list of Excel files, maps old external link paths to new locations, updates each workbook's Worksheets.ExternalLinks DataSource, recalculates all formulas, and saves the changes. Ideal for automating link migration across many workbooks.
    class Program
    {
        static void Main()
        {
            // List of workbook files to process
            string[] workbookFiles = new string[]
            {
                @"C:\Workbooks\Book1.xlsx",
                @"C:\Workbooks\Book2.xlsx",
                // add more files as needed
            };

            // Mapping of old external link file names (or full paths) to new locations
            var linkUpdates = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // key = existing external link data source, value = new data source
                { @"C:\OldLinks\External1.xlsx", @"D:\NewLinks\External1_v2.xlsx" },
                { @"C:\OldLinks\External2.xlsx", @"D:\NewLinks\External2_v2.xlsx" }
                // add more mappings as needed
            };

            foreach (var filePath in workbookFiles)
            {
                // Verify that the workbook file exists before attempting to load it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (lifecycle rule: load)
                    Workbook wb = new Workbook(filePath);

                    // Update each external link if it matches an entry in the mapping
                    foreach (ExternalLink extLink in wb.Worksheets.ExternalLinks)
                    {
                        // extLink.DataSource holds the current external file reference
                        if (linkUpdates.TryGetValue(extLink.DataSource, out string newSource))
                        {
                            // Update the external link to point to the new location
                            extLink.DataSource = newSource;
                        }
                    }

                    // Recalculate formulas so that values reflect the new external data
                    wb.CalculateFormula();

                    // Save the workbook (lifecycle rule: save)
                    wb.Save(filePath); // overwrites the original file; change path if a different output is required
                }
                catch (Exception ex)
                {
                    // Log the error and continue with the next workbook
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch external link update completed.");
        }
    }
}
