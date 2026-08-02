// Title: Add Aspose.Cells via NuGet and Merge Excel Workbooks with CellsHelper.MergeFiles in C#
// Description: A console example that shows how to install the Aspose.Cells library using the dotnet CLI, create two simple workbooks, merge them into one file with CellsHelper.MergeFiles (using a temporary cache file), display each sheet's A1 value, and clean up the source and cache files.
// Keywords: Aspose.Cells | NuGet package | C# | .NET | CellsHelper.MergeFiles | Excel workbook merge | sample workbook creation | temporary cache file | delete files after merge | console application
// Common Searches: install Aspose.Cells NuGet and merge workbooks | C# merge multiple xlsx files Aspose.Cells | CellsHelper.MergeFiles example | how to use temporary cache file with Aspose.Cells | combine Excel files in .NET console app
// Developer Intent: Install the Aspose.Cells NuGet package and use CellsHelper.MergeFiles to combine several Excel files into a single workbook.
// Use Cases: Consolidate daily report spreadsheets into one summary workbook. | Automate merging of generated Excel files in a CI/CD pipeline. | Create a quick prototype that merges temporary workbooks for testing purposes.
// AI Prompts: Generate a C# method that accepts a list of .xlsx paths and merges them with Aspose.Cells CellsHelper.MergeFiles, including validation for missing or corrupted files. | Write a PowerShell script that installs Aspose.Cells via dotnet CLI, builds the project, and runs the console app that merges Excel workbooks.

using System;
using System.IO;
using Aspose.Cells;

// A console example that shows how to install the Aspose.Cells library using the dotnet CLI, create two simple workbooks, merge them into one file with CellsHelper.MergeFiles (using a temporary cache file), display each sheet's A1 value, and clean up the source and cache files.
class Program
{
    static void Main()
    {
        // Install Aspose.Cells via NuGet before building:
        // dotnet add package Aspose.Cells

        // Prepare sample workbooks to be merged
        string[] filesToMerge = { "File1.xlsx", "File2.xlsx" };
        CreateSampleWorkbook(filesToMerge[0], "Content from File 1");
        CreateSampleWorkbook(filesToMerge[1], "Content from File 2");

        // Temporary cache file required by CellsHelper.MergeFiles
        string cacheFile = "CacheFile.tmp";

        // Destination merged workbook
        string mergedFile = "MergedOutput.xlsx";

        // Merge the files using the CellsHelper.MergeFiles method (provided rule)
        CellsHelper.MergeFiles(filesToMerge, cacheFile, mergedFile);

        // Verify the merged result
        Workbook mergedWorkbook = new Workbook(mergedFile);
        Console.WriteLine("Merged workbook contains the following sheets and values:");
        for (int i = 0; i < mergedWorkbook.Worksheets.Count; i++)
        {
            Worksheet sheet = mergedWorkbook.Worksheets[i];
            Console.WriteLine($"Sheet {i + 1} - A1: {sheet.Cells["A1"].StringValue}");
        }

        // Clean up temporary files
        foreach (string file in filesToMerge)
        {
            if (File.Exists(file)) File.Delete(file);
        }
        if (File.Exists(cacheFile)) File.Delete(cacheFile);
    }

    // Helper method to create a simple workbook with a single cell value
    static void CreateSampleWorkbook(string path, string value)
    {
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue(value);
        wb.Save(path);
    }
}
