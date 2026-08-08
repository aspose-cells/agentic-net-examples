// Title: Log merged workbook size after CellsHelper.MergeFiles in Aspose.Cells for .NET
// Description: Creates two temporary Excel files, merges them with CellsHelper.MergeFiles, saves the result, and writes the merged workbook's byte size to the console while cleaning up temporary files.
// Keywords: Aspose.Cells | CellsHelper.MergeFiles | C# | .NET | merge Excel workbooks | log file size | merged workbook size | FileInfo length | storage impact | Excel file size measurement
// Common Searches: How to get size of merged workbook using Aspose.Cells | Log file size after CellsHelper.MergeFiles C# | Measure storage impact of Excel merge Aspose | Retrieve merged Excel file byte count .NET | Aspose.Cells merge files and check output size
// Developer Intent: The developer needs to combine multiple Excel workbooks and record the resulting file size to monitor storage usage or enforce size limits.
// Use Cases: Automate daily report consolidation and capture the combined file size for quota monitoring. | Merge client‑provided spreadsheets and verify the final workbook stays within upload limits. | Aggregate data extracts into a single workbook and log its size for audit trails.
// AI Prompts: Generate C# code that merges a list of Excel files with Aspose.Cells and prints the merged file size in bytes. | Show how to handle exceptions, clean temporary files, and log the merged workbook size using Aspose.Cells. | Explain how to format the merged workbook size in kilobytes or megabytes for readable console output.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeAndLogSize
{
    // Creates two temporary Excel files, merges them with CellsHelper.MergeFiles, saves the result, and writes the merged workbook's byte size to the console while cleaning up temporary files.
    class Program
    {
        static void Main()
        {
            // Prepare temporary files to merge
            string[] filesToMerge = new string[2];
            filesToMerge[0] = "TempFile1.xlsx";
            filesToMerge[1] = "TempFile2.xlsx";

            // Create first workbook and save
            Workbook wb1 = new Workbook();
            wb1.Worksheets[0].Cells["A1"].PutValue("Data from first file");
            wb1.Save(filesToMerge[0]);

            // Create second workbook and save
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Cells["A1"].PutValue("Data from second file");
            wb2.Save(filesToMerge[1]);

            // Define cache and output files
            string cacheFile = "MergeCache.tmp";
            string outputFile = "MergedResult.xlsx";

            try
            {
                // Merge the temporary files into a single workbook
                CellsHelper.MergeFiles(filesToMerge, cacheFile, outputFile);

                // Log the size of the merged workbook
                FileInfo mergedInfo = new FileInfo(outputFile);
                Console.WriteLine($"Merged workbook saved to '{outputFile}'.");
                Console.WriteLine($"File size: {mergedInfo.Length} bytes");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merge: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files
                foreach (string path in filesToMerge)
                {
                    if (File.Exists(path)) File.Delete(path);
                }
                if (File.Exists(cacheFile)) File.Delete(cacheFile);
            }
        }
    }
}
