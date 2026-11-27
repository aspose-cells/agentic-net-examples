// Title: Clone an Excel workbook, renumber worksheet TabIds, delete empty sheets, and save the optimized copy with Aspose.Cells for .NET
// Common Searches: asp.net clone excel workbook and delete blank worksheets using Aspose.Cells | how to assign sequential TabId values to worksheets in C# with Aspose.Cells | save a cleaned copy of an Excel file after removing empty sheets in Aspose.Cells .NET | copy workbook while preserving macros and renumbering tabs in C# Aspose.Cells
// Tags: clone workbook Aspose.Cells C# | delete empty worksheets Aspose.Cells | assign sequential TabId Aspose.Cells | save workbook as Xlsx Aspose.Cells | copy workbook preserving macros Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookCloneAndOptimize
{
    // The program loads a source Excel file, clones its entire content into a new Workbook, removes any worksheets that have no data, assigns new sequential TabId values to the remaining sheets, and saves the cleaned workbook as an optimized XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths
            string inputPath;
            string outputPath;

            if (args.Length >= 2)
            {
                inputPath = args[0];
                outputPath = args[1];
            }
            else
            {
                // Default file names in the executable directory
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                inputPath = Path.Combine(exeDir, "SourceWorkbook.xlsx");
                outputPath = Path.Combine(exeDir, "OptimizedCopy.xlsx");
                Console.WriteLine("Usage: WorkbookCloneAndOptimize <inputPath> <outputPath>");
                Console.WriteLine($"No arguments supplied – using defaults:");
                Console.WriteLine($"  Input : {inputPath}");
                Console.WriteLine($"  Output: {outputPath}");
            }

            // Validate input file existence
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Error: Input file does not exist – \"{inputPath}\"");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook \"{inputPath}\" with {sourceWorkbook.Worksheets.Count} worksheet(s).");

                // Create a new workbook that will hold the cloned data
                Workbook clonedWorkbook = new Workbook();

                // Copy all content from source to the new workbook
                // Using Copy with default options to preserve everything (including macros if present)
                clonedWorkbook.Copy(sourceWorkbook);
                Console.WriteLine("Workbook content cloned successfully.");

                // Assign new TabIds (sequential starting from 1) and remove empty worksheets
                WorksheetCollection sheets = clonedWorkbook.Worksheets;
                int nextTabId = 1;

                // Iterate backwards so that removal does not affect the loop index
                for (int i = sheets.Count - 1; i >= 0; i--)
                {
                    Worksheet ws = sheets[i];

                    // Determine if the worksheet is empty:
                    // A worksheet is considered empty when it has no used rows and no used columns.
                    bool isEmpty = ws.Cells.MaxDataRow < 0 && ws.Cells.MaxDataColumn < 0;

                    if (isEmpty)
                    {
                        Console.WriteLine($"Removing empty worksheet \"{ws.Name}\" (index {i}).");
                        sheets.RemoveAt(i);
                        continue; // Skip TabId assignment for removed sheet
                    }

                    // Assign a new TabId
                    ws.TabId = nextTabId++;
                    Console.WriteLine($"Worksheet \"{ws.Name}\" assigned new TabId {ws.TabId}.");
                }

                Console.WriteLine($"After cleanup, workbook contains {clonedWorkbook.Worksheets.Count} worksheet(s).");

                // Save the optimized copy
                clonedWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Optimized workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An unexpected error occurred:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
