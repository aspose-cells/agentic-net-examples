using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.WebExtensions; // Example of additional namespace if needed

// Ensure the Aspose.Cells library is added to the project via NuGet:
// Install-Package Aspose.Cells

namespace AsposeCellsMergeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Prepare temporary files to demonstrate merging
            string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };
            CreateSampleWorkbook(sourceFiles[0], "Data from file 1");
            CreateSampleWorkbook(sourceFiles[1], "Data from file 2");

            // Define temporary cache file and destination merged file
            string cacheFile = "MergeCache.tmp";
            string mergedFile = "MergedResult.xlsx";

            try
            {
                // Merge the source files into a single workbook using the provided CellsHelper.MergeFiles method
                CellsHelper.MergeFiles(sourceFiles, cacheFile, mergedFile);

                // Load the merged workbook to verify the result
                Workbook mergedWorkbook = new Workbook(mergedFile);
                Console.WriteLine("Merged workbook created successfully.");
                Console.WriteLine("Sheet count: " + mergedWorkbook.Worksheets.Count);
                Console.WriteLine("First sheet, cell A1: " + mergedWorkbook.Worksheets[0].Cells["A1"].StringValue);
                Console.WriteLine("Second sheet, cell A1: " + mergedWorkbook.Worksheets[1].Cells["A1"].StringValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during merge: " + ex.Message);
            }
            finally
            {
                // Clean up temporary files
                foreach (var file in sourceFiles)
                {
                    if (File.Exists(file)) File.Delete(file);
                }
                if (File.Exists(cacheFile)) File.Delete(cacheFile);
                // Optionally keep the merged file for inspection
            }
        }

        // Helper method to create a simple workbook with a single cell value
        private static void CreateSampleWorkbook(string filePath, string cellValue)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue(cellValue);
            wb.Save(filePath);
        }
    }
}