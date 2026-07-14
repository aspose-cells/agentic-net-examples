// Add Aspose.Cells to the project via NuGet before running this code:
//   Install-Package Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Prepare sample workbooks that will be merged
        string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };
        CreateSampleWorkbook(sourceFiles[0], "Data from file 1");
        CreateSampleWorkbook(sourceFiles[1], "Data from file 2");

        // Temporary cached file required by CellsHelper.MergeFiles
        string cacheFile = "mergeCache.tmp";
        string destFile = "MergedResult.xlsx";

        // Merge the source files into a single workbook
        CellsHelper.MergeFiles(sourceFiles, cacheFile, destFile);

        // Load the merged workbook to verify the result
        Workbook mergedWorkbook = new Workbook(destFile);
        Console.WriteLine($"Merged workbook contains {mergedWorkbook.Worksheets.Count} worksheets.");
        Console.WriteLine($"First sheet A1: {mergedWorkbook.Worksheets[0].Cells["A1"].StringValue}");
        Console.WriteLine($"Second sheet A1: {mergedWorkbook.Worksheets[1].Cells["A1"].StringValue}");

        // Clean up temporary files
        foreach (string file in sourceFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
        if (File.Exists(cacheFile))
            File.Delete(cacheFile);
    }

    // Helper method to create a simple workbook with a single value
    static void CreateSampleWorkbook(string path, string cellValue)
    {
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue(cellValue);
        wb.Save(path);
    }
}