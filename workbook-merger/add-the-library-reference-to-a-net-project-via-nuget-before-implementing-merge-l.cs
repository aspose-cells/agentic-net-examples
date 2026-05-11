using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Add Aspose.Cells to the project via NuGet:
        //   Install-Package Aspose.Cells
        // The library reference is required for the merge logic below.

        // Prepare sample workbooks that will be merged.
        string[] files = new string[2];
        files[0] = "Part1.xlsx";
        files[1] = "Part2.xlsx";

        // Create first workbook and add some data.
        Workbook wb1 = new Workbook();
        wb1.Worksheets[0].Cells["A1"].PutValue("Data from first file");
        wb1.Save(files[0]);

        // Create second workbook and add some data.
        Workbook wb2 = new Workbook();
        wb2.Worksheets[0].Cells["A1"].PutValue("Data from second file");
        wb2.Save(files[1]);

        // Temporary cache file required by CellsHelper.MergeFiles.
        string cacheFile = "mergeCache.tmp";
        string resultFile = "MergedResult.xlsx";

        // Merge the two workbooks into a single file.
        CellsHelper.MergeFiles(files, cacheFile, resultFile);

        // Load the merged workbook to verify the result.
        Workbook merged = new Workbook(resultFile);
        Console.WriteLine("Merged workbook contains:");
        Console.WriteLine("Sheet0 A1: " + merged.Worksheets[0].Cells["A1"].StringValue);
        Console.WriteLine("Sheet1 A1: " + merged.Worksheets[1].Cells["A1"].StringValue);

        // Clean up temporary files.
        File.Delete(files[0]);
        File.Delete(files[1]);
        File.Delete(cacheFile);
    }
}