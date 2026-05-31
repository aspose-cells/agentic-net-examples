using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeAndLogSize
{
    class Program
    {
        static void Main()
        {
            // Prepare temporary files to be merged
            string[] filesToMerge = new string[2];
            filesToMerge[0] = "TempFile1.xlsx";
            filesToMerge[1] = "TempFile2.xlsx";

            // Create first workbook and save it
            Workbook wb1 = new Workbook();
            wb1.Worksheets[0].Cells["A1"].PutValue("Data from first workbook");
            wb1.Save(filesToMerge[0]); // using provided Save(string) method

            // Create second workbook and save it
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Cells["A1"].PutValue("Data from second workbook");
            wb2.Save(filesToMerge[1]); // using provided Save(string) method

            // Define temporary cache file and output merged file
            string cacheFile = "MergeCache.tmp";
            string mergedOutput = "MergedResult.xlsx";

            // Merge the files using the provided CellsHelper.MergeFiles method
            CellsHelper.MergeFiles(filesToMerge, cacheFile, mergedOutput);

            // After saving, obtain the file size
            FileInfo mergedInfo = new FileInfo(mergedOutput);
            long fileSizeInBytes = mergedInfo.Length;

            // Log the size (in kilobytes for readability)
            Console.WriteLine($"Merged workbook saved to '{mergedOutput}'.");
            Console.WriteLine($"File size: {fileSizeInBytes} bytes ({fileSizeInBytes / 1024.0:F2} KB).");

            // Clean up temporary files
            foreach (string path in filesToMerge)
            {
                if (File.Exists(path)) File.Delete(path);
            }
            if (File.Exists(cacheFile)) File.Delete(cacheFile);
            // Optionally keep the merged file; delete if not needed
            // if (File.Exists(mergedOutput)) File.Delete(mergedOutput);
        }
    }
}