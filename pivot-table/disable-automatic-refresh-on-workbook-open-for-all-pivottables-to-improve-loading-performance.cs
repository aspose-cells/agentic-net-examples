// Title: Disable PivotTable Auto‑Refresh on Workbook Open with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, iterates all worksheets, sets each PivotTable's RefreshDataOnOpeningFile property to false, and saves the workbook, preventing automatic refresh and speeding up load time.
// Keywords: Aspose.Cells | C# | PivotTable | RefreshDataOnOpeningFile | disable auto refresh | Excel load performance | prevent pivot refresh | Aspose.Cells .NET example
// Common Searches: Aspose.Cells disable pivot auto refresh | C# set RefreshDataOnOpeningFile false | prevent pivot tables from refreshing on open | improve Excel load time Aspose.Cells | how to turn off pivot refresh in .NET
// Developer Intent: Turn off automatic refresh for every PivotTable when the workbook is opened.
// Use Cases: Open a large workbook with many PivotTables without the overhead of data refresh. | Create a template that keeps PivotTable data static until the user manually refreshes it. | Batch‑process multiple files to ensure none of their PivotTables auto‑refresh on load.
// AI Prompts: Generate C# code using Aspose.Cells that disables RefreshDataOnOpeningFile for all PivotTables in a workbook and saves the file. | Show how to safely handle missing input files while iterating worksheets and setting the auto‑refresh flag to false. | Explain how to modify the sample to disable auto‑refresh only for PivotTables whose names match a given pattern.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads an Excel file, iterates all worksheets, sets each PivotTable's RefreshDataOnOpeningFile property to false, and saves the workbook, preventing automatic refresh and speeding up load time.
    public class DisablePivotAutoRefresh
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and their pivot tables
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        // Disable automatic refresh when the workbook is opened
                        pivot.RefreshDataOnOpeningFile = false;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DisablePivotAutoRefresh.Run();
        }
    }
}
