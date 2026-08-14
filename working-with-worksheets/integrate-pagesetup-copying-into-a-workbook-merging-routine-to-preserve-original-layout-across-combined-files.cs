// Title: Merge Excel workbooks and retain page‑setup settings with Aspose.Cells for .NET
// Description: Demonstrates how to combine multiple .xlsx files into a single workbook using Workbook.Combine, then copy each source worksheet's PageSetup to the merged sheet with Worksheet.PageSetup.Copy and CopyOptions, preserving print layout and margins.
// Keywords: Aspose.Cells merge workbooks | copy page setup Aspose.Cells | Workbook.Combine page setup | preserve worksheet print settings | C# Excel workbook consolidation | Aspose.Cells PageSetup.Copy example | merge .xlsx files .NET
// Common Searches: Aspose.Cells copy page setup after combine | preserve print settings when merging Excel files C# | how to keep worksheet layout after Workbook.Combine | merge multiple Excel workbooks without losing page setup | Aspose.Cells combine workbooks with page orientation
// Developer Intent: The developer needs to merge several Excel workbooks into one while keeping each worksheet’s original page‑setup configuration (margins, orientation, scaling, headers/footers).
// Use Cases: Consolidate monthly reports into a master workbook while retaining each sheet’s print layout. | Create a single printable file from departmental templates without losing paper size or header/footer settings. | Automate generation of a combined workbook for distribution where every source sheet must keep its exact page formatting.
// AI Prompts: Write C# code that merges a list of Excel files using Aspose.Cells and copies the PageSetup of each source worksheet to the corresponding destination worksheet. | Explain the role of CopyOptions when using Worksheet.PageSetup.Copy after a Workbook.Combine operation. | Provide robust error‑handling patterns for loading multiple workbooks, merging them, and preserving page‑setup settings with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to combine multiple .xlsx files into a single workbook using Workbook.Combine, then copy each source worksheet's PageSetup to the merged sheet with Worksheet.PageSetup.Copy and CopyOptions, preserving print layout and margins.
    public class MergeWorkbooksWithPageSetup
    {
        public static void Run()
        {
            // Destination workbook that will hold the merged result
            Workbook destWorkbook = null;
            try
            {
                destWorkbook = new Workbook();

                // List of source workbook file paths to be merged
                string[] sourceFiles = new string[]
                {
                    "Source1.xlsx",
                    "Source2.xlsx",
                    "Source3.xlsx"
                };

                // Iterate through each source workbook
                foreach (string srcPath in sourceFiles)
                {
                    // Verify source file exists to avoid FileNotFoundException
                    if (!File.Exists(srcPath))
                    {
                        Console.WriteLine($"Source file not found: {srcPath}. Skipping.");
                        continue;
                    }

                    try
                    {
                        // Load the source workbook
                        using (Workbook srcWorkbook = new Workbook(srcPath))
                        {
                            // Record the number of worksheets before combining
                            int beforeCombineCount = destWorkbook.Worksheets.Count;

                            // Combine the source workbook into the destination workbook
                            destWorkbook.Combine(srcWorkbook);

                            // Record the number of worksheets after combining
                            int afterCombineCount = destWorkbook.Worksheets.Count;

                            // Copy PageSetup from each source worksheet to its corresponding newly added worksheet
                            for (int i = beforeCombineCount; i < afterCombineCount; i++)
                            {
                                // Index of the worksheet in the source workbook that matches the newly added one
                                int srcIndex = i - beforeCombineCount;

                                Worksheet destSheet = destWorkbook.Worksheets[i];
                                Worksheet srcSheet = srcWorkbook.Worksheets[srcIndex];

                                // Copy page setup settings
                                destSheet.PageSetup.Copy(srcSheet.PageSetup, new CopyOptions());
                            }
                        } // srcWorkbook disposed here
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{srcPath}': {ex.Message}");
                    }
                }

                // Save the merged workbook with all page‑setup settings preserved
                string outputPath = "MergedWithPageSetup.xlsx";
                destWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbooks merged successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
            finally
            {
                // Clean up the destination workbook
                destWorkbook?.Dispose();
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MergeWorkbooksWithPageSetup.Run();
        }
    }
}
