using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ExternalLinkMigration
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook that needs to be inspected
            string sourceFile = "SourceWorkbook.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(sourceFile);

            // Collection to hold detected network external links
            List<ExternalLink> networkLinks = new List<ExternalLink>();

            // -----------------------------------------------------------------
            // 1. Detect external links defined in the workbook's ExternalLinks collection
            // -----------------------------------------------------------------
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                // DataSource contains the path used in the external reference
                string dataSource = link.DataSource;

                // UNC network paths start with "\\" (e.g., \\Server\Share\file.xlsx)
                if (dataSource.StartsWith(@"\\"))
                {
                    networkLinks.Add(link);
                }
            }

            // -----------------------------------------------------------------
            // 2. Detect formulas that embed network paths but are not captured
            //    in the ExternalLinks collection (e.g., indirect references)
            // -----------------------------------------------------------------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Simple heuristic: look for a UNC path inside the formula string
                        if (formula.Contains(@"\\"))
                        {
                            // Attempt to locate the corresponding ExternalLink; if not found, add a placeholder
                            bool alreadyTracked = false;
                            foreach (ExternalLink link in networkLinks)
                            {
                                if (formula.Contains(link.DataSource))
                                {
                                    alreadyTracked = true;
                                    break;
                                }
                            }

                            if (!alreadyTracked)
                            {
                                // Create a dummy ExternalLink object to represent this formula reference
                                // (Aspose.Cells does not allow creating ExternalLink directly, so we just note the path)
                                Console.WriteLine($"[Info] Formula in {sheet.Name}!{cell.Name} contains a network path not listed in ExternalLinks.");
                            }
                        }
                    }
                }
            }

            // -----------------------------------------------------------------
            // 3. Generate migration checklist for each detected network link
            // -----------------------------------------------------------------
            Console.WriteLine("\n=== Migration Checklist for Network External Links ===\n");

            foreach (ExternalLink link in networkLinks)
            {
                string originalPath = link.DataSource;
                Console.WriteLine($"External Link: {originalPath}");
                Console.WriteLine("  [ ] Verify that the network location is reachable.");
                Console.WriteLine("  [ ] Copy the external workbook to a local folder (e.g., C:\\LocalData\\).");
                Console.WriteLine("  [ ] Update the workbook to point to the new local copy.");

                // Example of updating the link (using OriginalDataSource which is settable)
                // Assume the local copy will be placed under C:\LocalData\ with the same file name
                string fileName = System.IO.Path.GetFileName(originalPath);
                string newLocalPath = System.IO.Path.Combine(@"C:\LocalData", fileName);

                // Update the stored data source
                link.OriginalDataSource = newLocalPath;

                Console.WriteLine($"  [ ] Updated link path set to: {newLocalPath}\n");
            }

            // -----------------------------------------------------------------
            // 4. Save the modified workbook (uses the provided save rule)
            // -----------------------------------------------------------------
            string outputFile = "SourceWorkbook_Migrated.xlsx";
            workbook.Save(outputFile);

            Console.WriteLine($"Workbook saved with updated external links to: {outputFile}");
        }
    }
}