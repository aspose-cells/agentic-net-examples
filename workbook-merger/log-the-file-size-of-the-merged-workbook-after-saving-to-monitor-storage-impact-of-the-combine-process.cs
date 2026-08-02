// Title: C# – Log Merged Workbook Size After CellsHelper.MergeFiles with Aspose.Cells
// Description: Creates two temporary Excel workbooks, merges them using Aspose.Cells CellsHelper.MergeFiles, then uses FileInfo to read the merged file's Length property and writes the size in bytes to the console. The example demonstrates how to monitor storage impact of the combine process and clean up temporary files.
// Keywords: Aspose.Cells | CellsHelper.MergeFiles | C# | merge Excel files | file size | log workbook size | storage impact | .NET | FileInfo.Length | Excel merge
// Common Searches: Aspose.Cells get merged file size C# | How to log size of merged Excel workbook .NET | CellsHelper.MergeFiles file size output | Check storage impact after merging Excel files with Aspose | C# code to merge workbooks and report size
// Developer Intent: Retrieve and record the byte size of the workbook produced by CellsHelper.MergeFiles to assess storage usage after merging.
// Use Cases: Validate that the combined workbook stays within storage quotas. | Automate size reporting in CI/CD pipelines after Excel merges. | Compare pre‑ and post‑merge file sizes to detect data duplication. | Trigger alerts when the merged file exceeds a predefined size limit. | Maintain an audit log of merged workbook sizes for compliance.
// AI Prompts: Generate C# code that merges a list of .xlsx files with Aspose.Cells and prints the resulting file size. | Show how to capture the merged workbook size in bytes, store it in a variable, and write it to a log file with proper exception handling. | Explain how to evaluate the storage impact of merging Excel workbooks and suggest Aspose.Cells options to reduce the final file size.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeAndLogSize
{
    // Creates two temporary Excel workbooks, merges them using Aspose.Cells CellsHelper.MergeFiles, then uses FileInfo to read the merged file's Length property and writes the size in bytes to the console. The example demonstrates how to monitor storage impact of the combine process and clean up temporary files.
    public class Program
    {
        public static void Main()
        {
            // Prepare temporary files to be merged
            string[] filesToMerge = new string[2];
            filesToMerge[0] = "TempFile1.xlsx";
            filesToMerge[1] = "TempFile2.xlsx";

            // Create first workbook, add data, and save
            Workbook wb1 = new Workbook();
            wb1.Worksheets[0].Cells["A1"].PutValue("Data from first workbook");
            wb1.Save(filesToMerge[0]);

            // Create second workbook, add data, and save
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Cells["A1"].PutValue("Data from second workbook");
            wb2.Save(filesToMerge[1]);

            // Define cached temporary file and final merged output file
            string cachedFile = "MergeCache.tmp";
            string mergedOutput = "MergedResult.xlsx";

            try
            {
                // Merge the temporary workbooks into a single file
                CellsHelper.MergeFiles(filesToMerge, cachedFile, mergedOutput);

                // Log the size of the merged workbook after saving
                FileInfo mergedInfo = new FileInfo(mergedOutput);
                Console.WriteLine($"Merged workbook saved to '{mergedOutput}'.");
                Console.WriteLine($"File size: {mergedInfo.Length} bytes");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merge operation: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files
                foreach (string path in filesToMerge)
                {
                    if (File.Exists(path)) File.Delete(path);
                }
                if (File.Exists(cachedFile)) File.Delete(cachedFile);
                // Optionally keep the merged file for further inspection
            }
        }
    }
}
