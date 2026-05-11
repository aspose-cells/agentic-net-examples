using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Paths of the source XLS files to be merged
            string[] filesToMerge = new string[2];
            filesToMerge[0] = "File1.xls";
            filesToMerge[1] = "File2.xls";

            // Create sample source workbooks (for demonstration purposes)
            // In a real scenario these files would already exist
            Workbook wb1 = new Workbook(); // create a new workbook
            wb1.Worksheets[0].Cells["A1"].PutValue("Content of File 1");
            wb1.Save(filesToMerge[0]); // save the first file

            Workbook wb2 = new Workbook(); // create another workbook
            wb2.Worksheets[0].Cells["A1"].PutValue("Content of File 2");
            wb2.Save(filesToMerge[1]); // save the second file

            // Temporary cached file required by MergeFiles
            string cachedFile = "CacheFile.tmp";

            // Destination file that will contain the merged result
            string outputFile = "MergedOutput.xls";

            try
            {
                // Merge the source files into the destination file
                // This uses the CellsHelper.MergeFiles method as defined in the API
                CellsHelper.MergeFiles(filesToMerge, cachedFile, outputFile);
                Console.WriteLine($"Files merged successfully. Output saved to: {outputFile}");

                // Load the merged workbook to verify the result
                Workbook mergedWorkbook = new Workbook(outputFile); // load the merged file
                Console.WriteLine("Merged workbook contains the following data:");
                Console.WriteLine($"Sheet 0, A1: {mergedWorkbook.Worksheets[0].Cells["A1"].StringValue}");
                if (mergedWorkbook.Worksheets.Count > 1)
                {
                    Console.WriteLine($"Sheet 1, A1: {mergedWorkbook.Worksheets[1].Cells["A1"].StringValue}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merging: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files used for the demo
                foreach (string path in filesToMerge)
                {
                    if (File.Exists(path)) File.Delete(path);
                }
                if (File.Exists(cachedFile)) File.Delete(cachedFile);
            }
        }
    }
}