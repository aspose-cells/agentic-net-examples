using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RefreshExternalDataConnectionsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the main workbook that contains external links
            string mainWorkbookPath = "MainWorkbook.xlsx";

            // Verify main workbook exists
            if (!File.Exists(mainWorkbookPath))
            {
                Console.WriteLine($"Main workbook not found: {mainWorkbookPath}");
                return;
            }

            // Paths to external workbooks referenced by the main workbook
            List<string> externalWorkbookPaths = new List<string>
            {
                "ExternalWorkbook1.xlsx",
                "ExternalWorkbook2.xlsx",
                // Add more external workbook file names as needed
            };

            // Load the main workbook inside a using block for proper disposal
            using (Workbook mainWorkbook = new Workbook(mainWorkbookPath))
            {
                // Iterate through each external workbook sequentially
                foreach (string externalPath in externalWorkbookPaths)
                {
                    // Skip missing external files
                    if (!File.Exists(externalPath))
                    {
                        Console.WriteLine($"External workbook not found and will be skipped: {externalPath}");
                        continue;
                    }

                    try
                    {
                        // Load the external workbook
                        using (Workbook externalWorkbook = new Workbook(externalPath))
                        {
                            // Refresh the external link in the main workbook
                            mainWorkbook.UpdateLinkedDataSource(new Workbook[] { externalWorkbook });

                            // Recalculate formulas to reflect the refreshed data
                            mainWorkbook.CalculateFormula();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing external workbook '{externalPath}': {ex.Message}");
                    }
                }

                // Save the updated main workbook
                string outputPath = "MainWorkbook_Refreshed.xlsx";
                try
                {
                    mainWorkbook.Save(outputPath);
                    Console.WriteLine($"External data connections refreshed and workbook saved to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
        }
    }
}