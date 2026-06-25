using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MergeLargeXlsFilesDemo
    {
        public static void Run()
        {
            // Files to be merged (using XLSX format which is always supported)
            string[] filesToMerge = { "File1.xlsx", "File2.xlsx" };

            // Create sample workbooks (large XLSX files) for demonstration
            try
            {
                Workbook wb1 = new Workbook();
                wb1.Worksheets[0].Cells["A1"].PutValue("Content of File 1");
                wb1.Save(filesToMerge[0], SaveFormat.Xlsx);

                Workbook wb2 = new Workbook();
                wb2.Worksheets[0].Cells["A1"].PutValue("Content of File 2");
                wb2.Save(filesToMerge[1], SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating sample files: {ex.Message}");
                return;
            }

            // Temporary cached file required by MergeFiles
            string cachedFile = "CacheFile.tmp";

            // Destination merged file
            string outputFile = "MergedOutput.xlsx";

            try
            {
                // Ensure source files exist before merging
                foreach (var f in filesToMerge)
                {
                    if (!File.Exists(f))
                        throw new FileNotFoundException($"Source file not found: {f}");
                }

                // Merge the large XLSX files
                CellsHelper.MergeFiles(filesToMerge, cachedFile, outputFile);
                Console.WriteLine($"Files merged successfully. Output saved to: {outputFile}");

                // Verify the merged workbook
                if (File.Exists(outputFile))
                {
                    Workbook mergedWorkbook = new Workbook(outputFile);
                    Console.WriteLine("Merged workbook contains the following data:");
                    Console.WriteLine($"Sheet 0, A1: {mergedWorkbook.Worksheets[0].Cells["A1"].StringValue}");
                    if (mergedWorkbook.Worksheets.Count > 1)
                    {
                        Console.WriteLine($"Sheet 1, A1: {mergedWorkbook.Worksheets[1].Cells["A1"].StringValue}");
                    }
                }
                else
                {
                    Console.WriteLine("Merged output file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merging: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files
                foreach (var file in filesToMerge)
                {
                    if (File.Exists(file)) File.Delete(file);
                }
                if (File.Exists(cachedFile)) File.Delete(cachedFile);
                // Optionally delete the merged output file
                // if (File.Exists(outputFile)) File.Delete(outputFile);
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            MergeLargeXlsFilesDemo.Run();
        }
    }
}