using System;
using System.IO;
using Aspose.Cells;

class WorkbookMerger
{
    /// <summary>
    /// Merges multiple Excel files into a single workbook while preserving data, styles, and formulas.
    /// </summary>
    /// <param name="inputFiles">Array of full paths to the source Excel files.</param>
    /// <param name="outputFile">Full path for the merged workbook.</param>
    public static void MergeFiles(string[] inputFiles, string outputFile)
    {
        // Temporary cache file required by CellsHelper.MergeFiles.
        string cacheFile = Path.Combine(Path.GetTempPath(), "AsposeMergeCache.tmp");

        try
        {
            // Perform the merge using the Aspose.Cells helper method.
            CellsHelper.MergeFiles(inputFiles, cacheFile, outputFile);
            Console.WriteLine($"Successfully merged {inputFiles.Length} files into '{outputFile}'.");
        }
        finally
        {
            // Clean up the temporary cache file.
            if (File.Exists(cacheFile))
                File.Delete(cacheFile);
        }
    }

    static void Main()
    {
        // Example source files (ensure these files exist or are created beforehand).
        string[] sourceFiles = new string[]
        {
            "File1.xlsx",
            "File2.xlsx",
            "File3.xlsx"
        };

        // Destination file for the merged result.
        string mergedFile = "CombinedOutput.xlsx";

        // Call the merge routine.
        MergeFiles(sourceFiles, mergedFile);
    }
}