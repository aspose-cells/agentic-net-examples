// Title: Install Aspose.Cells via NuGet and Merge Multiple Excel Workbooks with CellsHelper.MergeFiles in C#
// Description: Step‑by‑step example that shows how to add the Aspose.Cells NuGet package to a .NET project, create sample workbooks, and combine them into a single file using CellsHelper.MergeFiles. The code handles a temporary cache file, cleans up intermediate files, and verifies the merged sheets.
// Keywords: Aspose.Cells NuGet | dotnet add package Aspose.Cells | C# merge Excel workbooks | CellsHelper.MergeFiles example | combine .xlsx files .NET | Excel workbook consolidation | temporary cache file merge | Aspose.Cells tutorial | GitHub Aspose.Cells merge demo | US developers Excel automation
// Common Searches: how to install Aspose.Cells with NuGet before merging workbooks | C# code to merge multiple Excel files using CellsHelper.MergeFiles | sample program that creates and merges Excel workbooks Aspose.Cells | remove temporary cache file after merging Excel files .NET | Aspose.Cells workbook merge command line example
// Developer Intent: Add the Aspose.Cells NuGet package to a .NET project and use CellsHelper.MergeFiles to combine several Excel workbooks into one file.
// Use Cases: Generate individual monthly reports, then merge them into a yearly summary workbook. | Automate consolidation of data‑entry sheets from multiple sources into a master file. | Create temporary workbooks for data transformation, merge them, and delete the intermediates automatically.
// AI Prompts: Generate a PowerShell script that runs `dotnet add package Aspose.Cells` and executes the merge program. | Improve error handling for CellsHelper.MergeFiles to log missing source files and output detailed exception data. | Adapt the example to preserve original worksheet names and retain formatting when merging multiple workbooks.

using System;
using System.IO;
using Aspose.Cells;

namespace MergeDemo
{
    // Before compiling this code, add the Aspose.Cells NuGet package to the project:
    //   dotnet add package Aspose.Cells
    // or via Visual Studio's NuGet Package Manager.

    // Step‑by‑step example that shows how to add the Aspose.Cells NuGet package to a .NET project, create sample workbooks, and combine them into a single file using CellsHelper.MergeFiles. The code handles a temporary cache file, cleans up intermediate files, and verifies the merged sheets.
    class Program
    {
        static void Main()
        {
            // Prepare sample workbooks to demonstrate merging.
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };

            // Create first source workbook.
            Workbook wb1 = new Workbook();
            wb1.Worksheets[0].Cells["A1"].PutValue("Data from file 1");
            wb1.Save(sourceFiles[0]);

            // Create second source workbook.
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Cells["A1"].PutValue("Data from file 2");
            wb2.Save(sourceFiles[1]);

            // Temporary cache file required by CellsHelper.MergeFiles.
            string cacheFile = "mergeCache.tmp";

            // Destination file that will contain merged content.
            string destFile = "MergedResult.xlsx";

            try
            {
                // Merge the source workbooks into the destination workbook.
                // This uses the CellsHelper.MergeFiles method as defined in the Aspose.Cells API.
                CellsHelper.MergeFiles(sourceFiles, cacheFile, destFile);

                Console.WriteLine($"Merge completed successfully. Output saved to '{destFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merge: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files.
                foreach (string file in sourceFiles)
                {
                    if (File.Exists(file))
                        File.Delete(file);
                }

                if (File.Exists(cacheFile))
                    File.Delete(cacheFile);
            }

            // Optional: Verify the merged workbook content.
            Workbook merged = new Workbook(destFile);
            Console.WriteLine("Merged workbook sheets and first cell values:");
            for (int i = 0; i < merged.Worksheets.Count; i++)
            {
                string value = merged.Worksheets[i].Cells["A1"].StringValue;
                Console.WriteLine($"Sheet {i + 1}: {value}");
            }
        }
    }
}
