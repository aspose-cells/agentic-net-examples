using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            WorkbookMergeDemo.Run();
        }
    }

    public class WorkbookMergeDemo
    {
        public static void Run()
        {
            // Destination workbook – starts empty
            Workbook destinationWorkbook = new Workbook();

            // Array of source workbook file paths to be merged
            string[] sourceFiles = new string[]
            {
                "File1.xlsx",
                "File2.xlsx",
                "File3.xlsx"
            };

            // Merge each source workbook into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook from file
                Workbook sourceWorkbook = new Workbook(filePath);

                // Combine the source workbook data (worksheets, data, formulas, etc.)
                destinationWorkbook.Combine(sourceWorkbook);

                // Merge named styles from the source workbook to preserve style definitions
                destinationWorkbook.MergeNamedStyles(sourceWorkbook);
            }

            // Save the combined workbook to disk
            destinationWorkbook.Save("CombinedOutput.xlsx", SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // Alternative approach for merging large XLS files using CellsHelper.
            // This method merges only data, styles, and formulas and requires a
            // temporary cache file.
            // -----------------------------------------------------------------
            string cachedFile = "CacheFile.tmp";
            string mergedLargeFile = "MergedLargeOutput.xlsx";

            // CellsHelper.MergeFiles merges the same source files into a single file.
            // It is useful when dealing with many large workbooks.
            CellsHelper.MergeFiles(sourceFiles, cachedFile, mergedLargeFile);

            // Clean up the temporary cache file
            if (File.Exists(cachedFile))
            {
                File.Delete(cachedFile);
            }

            Console.WriteLine($"Merging completed. Files saved as 'CombinedOutput.xlsx' and '{mergedLargeFile}'.");
        }
    }
}