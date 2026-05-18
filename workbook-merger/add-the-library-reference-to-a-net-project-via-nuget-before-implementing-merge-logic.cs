using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // NOTE: Before running this code, add the Aspose.Cells via NuGet:
            //   Install-Package Aspose.Cells
            // This ensures the required library is referenced in the project.

            // Prepare sample workbooks to demonstrate merge logic.
            string[] filesToMerge = new string[2];
            filesToMerge[0] = "File1.xlsx";
            filesToMerge[1] = "File2.xlsx";

            // Create first workbook and add some data.
            Workbook wb1 = new Workbook();
            wb1.Worksheets[0].Cells["A1"].PutValue("Content from File 1");
            wb1.Save(filesToMerge[0]);

            // Create second workbook and add some data.
            Workbook wb2 = new Workbook();
            wb2.Worksheets[0].Cells["A1"].PutValue("Content from File 2");
            wb2.Save(filesToMerge[1]);

            // Define a temporary cached file required by CellsHelper.MergeFiles.
            string cachedFile = "CacheFile.tmp";

            // Define the destination file that will contain the merged result.
            string outputFile = "MergedOutput.xlsx";

            try
            {
                // Merge the source files into the destination file.
                // This uses the CellsHelper.MergeFiles method as defined in the API.
                CellsHelper.MergeFiles(filesToMerge, cachedFile, outputFile);

                Console.WriteLine($"Files merged successfully. Output saved to: {outputFile}");

                // Verify the merge by loading the resulting workbook.
                Workbook mergedWorkbook = new Workbook(outputFile);
                Console.WriteLine("Merged workbook content:");
                Console.WriteLine($"Sheet1!A1 = {mergedWorkbook.Worksheets[0].Cells["A1"].StringValue}");
                if (mergedWorkbook.Worksheets.Count > 1)
                {
                    Console.WriteLine($"Sheet2!A1 = {mergedWorkbook.Worksheets[1].Cells["A1"].StringValue}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merge: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files used for the demo.
                foreach (var file in filesToMerge)
                {
                    if (File.Exists(file))
                        File.Delete(file);
                }

                if (File.Exists(cachedFile))
                    File.Delete(cachedFile);
            }
        }
    }
}