// Title: Copy Selected Worksheets from Multiple Excel Workbooks with Aspose.Cells Worksheet.Copy (C#)
// Description: C# sample that loads several source .xlsx files, selects specific sheets, creates an empty workbook, adds matching worksheets, copies content and formatting via Worksheet.Copy, and saves the combined file while handling missing files or sheets.
// Keywords: Aspose.Cells copy worksheets C# | Worksheet.Copy example | merge specific sheets Aspose | combine Excel workbooks programmatically | copy sheet formatting Aspose.Cells
// Common Searches: Aspose.Cells copy specific worksheets from multiple workbooks | C# Worksheet.Copy preserve formatting | merge selected Excel sheets using Aspose | how to combine sheets from different files with Aspose.Cells | copy sheet by name to another workbook C#
// Developer Intent: Copy chosen worksheets from several source workbooks into one target workbook, preserving data, formulas, styles, and merged cells.
// Use Cases: Consolidate "Sheet1" and "Data" from Source1.xlsx with "Report" from Source2.xlsx into a single file. | Build a master report that aggregates only required sheets from departmental workbooks. | Automate archival of specific worksheets without opening Excel manually.
// AI Prompts: Generate a C# snippet that copies a list of worksheet names from multiple source workbooks into a new workbook using Aspose.Cells, with graceful handling of missing files or sheets. | Explain how Worksheet.Copy retains formulas, cell styles, and merged ranges during a merge operation. | Suggest ways to prevent duplicate sheet names and to define the order of copied worksheets in the destination workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    // C# sample that loads several source .xlsx files, selects specific sheets, creates an empty workbook, adds matching worksheets, copies content and formatting via Worksheet.Copy, and saves the combined file while handling missing files or sheets.
    class Program
    {
        static void Main()
        {
            // Paths of source workbooks
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };

            // Define which worksheets to copy from each source workbook
            // The outer array corresponds to sourceFiles order
            string[][] sheetsToCopy = {
                new string[] { "Sheet1", "Data" },   // from Source1.xlsx
                new string[] { "Report" }           // from Source2.xlsx
            };

            // Create the target workbook (initially empty) and remove the default worksheet
            Workbook targetWorkbook = new Workbook();
            targetWorkbook.Worksheets.Clear();

            try
            {
                // Iterate over each source workbook
                for (int i = 0; i < sourceFiles.Length; i++)
                {
                    string sourcePath = sourceFiles[i];

                    // Verify source file exists
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"Source file not found: {sourcePath}. Skipping.");
                        continue;
                    }

                    // Load the source workbook
                    Workbook sourceWorkbook = new Workbook(sourcePath);

                    // Ensure there is a corresponding sheet list
                    if (i >= sheetsToCopy.Length)
                    {
                        Console.WriteLine($"No sheet list defined for {sourcePath}. Skipping.");
                        continue;
                    }

                    // Iterate over the worksheets that need to be copied from this source
                    foreach (string sheetName in sheetsToCopy[i])
                    {
                        // Get the source worksheet by name
                        Worksheet sourceSheet = sourceWorkbook.Worksheets[sheetName];

                        if (sourceSheet == null)
                        {
                            Console.WriteLine($"Worksheet \"{sheetName}\" not found in {sourcePath}. Skipping.");
                            continue;
                        }

                        // Add a new worksheet to the target workbook with the same name
                        Worksheet destSheet = targetWorkbook.Worksheets.Add(sourceSheet.Name);

                        // Copy the contents and formatting from the source worksheet to the destination worksheet
                        destSheet.Copy(sourceSheet);
                    }
                }

                // Save the merged workbook containing the copied worksheets
                string outputPath = "MergedOutput.xlsx";
                targetWorkbook.Save(outputPath);
                Console.WriteLine($"Merged workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
