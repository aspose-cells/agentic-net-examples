using System;
using System.IO;
using Aspose.Cells;

namespace MergeWorkbookSizeLogger
{
    class Program
    {
        static void Main()
        {
            // Prepare temporary files to merge
            string[] filesToMerge = new string[2];
            filesToMerge[0] = "TempFile1.xlsx";
            filesToMerge[1] = "TempFile2.xlsx";

            // Create first workbook and add sample data
            Workbook wb1 = new Workbook();
            wb1.Worksheets[0].Cells["A1"].PutValue("Data from workbook 1");
            wb1.Save(filesToMerge[0]); // using provided Save(string) method

            // Create second workbook and add sample data
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Cells["A1"].PutValue("Data from workbook 2");
            wb2.Save(filesToMerge[1]); // using provided Save(string) method

            // Define cache file and output merged file
            string cacheFile = "MergeCache.tmp";
            string mergedFile = "MergedOutput.xlsx";

            try
            {
                // Merge the temporary workbooks into a single file
                CellsHelper.MergeFiles(filesToMerge, cacheFile, mergedFile); // using provided MergeFiles method

                // Log the size of the merged workbook
                long fileSizeBytes = new FileInfo(mergedFile).Length;
                Console.WriteLine($"Merged workbook saved to '{mergedFile}'. Size: {fileSizeBytes} bytes.");
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
                    if (File.Exists(path))
                        File.Delete(path);
                }

                if (File.Exists(cacheFile))
                    File.Delete(cacheFile);
            }
        }
    }
}