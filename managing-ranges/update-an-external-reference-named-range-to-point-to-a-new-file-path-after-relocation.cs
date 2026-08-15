// Title: C# – Update External Named Range Path with Aspose.Cells for .NET
// Description: Loads a workbook, scans all defined names, identifies external links, updates the ExternalLink.DataSource and OriginalDataSource to a new file location, rewrites the RefersTo formula, and saves the workbook with corrected external references.
// Keywords: Aspose.Cells external link update | C# update named range path | Excel external reference relocation | Modify DataSource Aspose.Cells | RefersTo formula path replace | .NET workbook external link | Update ExternalLink DataSource | Change external file source Excel
// Common Searches: Aspose.Cells change external named range file path C# | Update external link path for defined names in Excel using .NET | Programmatically fix broken external references after moving source workbook | Replace external data source in RefersTo formula Aspose.Cells | How to rewrite external link paths for named ranges
// Developer Intent: Rewrite the file path of an external named‑range reference so the workbook points to the new location.
// Use Cases: Repair broken external links after moving source files to a new directory. | Batch‑process multiple workbooks to re‑map external data sources during a folder restructure. | Integrate path‑update logic into CI/CD pipelines to ensure Excel reports reference the correct data files before deployment.
// AI Prompts: Write C# code with Aspose.Cells that updates all external named‑range links to a new folder path. | Explain how to locate ExternalLink objects and modify their DataSource and RefersTo values for defined names. | Provide robust error‑handling patterns when changing external references in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;

// Loads a workbook, scans all defined names, identifies external links, updates the ExternalLink.DataSource and OriginalDataSource to a new file location, rewrites the RefersTo formula, and saves the workbook with corrected external references.
class UpdateExternalReferenceNamedRange
{
    static void Main()
    {
        try
        {
            const string originalPath = "OriginalWorkbook.xlsx";
            const string updatedPath = "UpdatedWorkbook.xlsx";
            const string newExternalPath = @"D:\NewFolder\ExternalData.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Source workbook not found: {originalPath}");
                return;
            }

            // Load the workbook that contains the external reference named range
            Workbook workbook = new Workbook(originalPath);

            // Ensure there are defined names to process
            if (workbook.Worksheets.Names == null || workbook.Worksheets.Names.Count == 0)
            {
                Console.WriteLine("No defined names found in the workbook.");
                return;
            }

            // Iterate through all defined names in the workbook
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                // Get all referred areas of the defined name (including external links)
                ReferredArea[] areas = definedName.GetReferredAreas(true);
                if (areas == null) continue;

                foreach (ReferredArea area in areas)
                {
                    // Process only external links
                    if (area.IsExternalLink)
                    {
                        // Search for the matching ExternalLink object
                        for (int i = 0; i < workbook.Worksheets.ExternalLinks.Count; i++)
                        {
                            ExternalLink extLink = workbook.Worksheets.ExternalLinks[i];

                            // Compare external file names (case‑insensitive)
                            if (string.Equals(extLink.DataSource, area.ExternalFileName, StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(extLink.OriginalDataSource, area.ExternalFileName, StringComparison.OrdinalIgnoreCase))
                            {
                                // Update the data source paths
                                extLink.DataSource = newExternalPath;
                                extLink.OriginalDataSource = newExternalPath;

                                // Update the RefersTo formula text to reflect the new path
                                string oldRefersTo = definedName.RefersTo;
                                if (!string.IsNullOrEmpty(oldRefersTo))
                                {
                                    string updatedRefersTo = oldRefersTo.Replace(area.ExternalFileName, newExternalPath);
                                    definedName.RefersTo = updatedRefersTo;
                                }

                                Console.WriteLine($"Updated external link for a named range to '{newExternalPath}'.");
                                break; // Exit the inner loop once the matching link is updated
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(updatedPath);
            Console.WriteLine($"Workbook saved as '{updatedPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
