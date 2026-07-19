// Title: Merge Multiple Excel Workbooks with Aspose.Cells in .NET – NuGet Installation & CellsHelper.MergeFiles Example
// Description: Step‑by‑step guide that shows how to add the Aspose.Cells library to a .NET project using the dotnet CLI, create sample workbooks, and combine them into a single file with CellsHelper.MergeFiles. The example loads the merged workbook, prints worksheet count, and displays A1 values for verification.
// Keywords: Aspose.Cells NuGet | C# merge Excel files | CellsHelper.MergeFiles | combine .xlsx workbooks .NET | sample Aspose.Cells code | Excel workbook consolidation | dotnet add package Aspose.Cells | verify merged workbook
// Common Searches: how to install Aspose.Cells via NuGet | C# merge multiple Excel files Aspose.Cells | CellsHelper.MergeFiles example | combine .xlsx files programmatically .NET | sample code for merging workbooks Aspose
// Developer Intent: Add the Aspose.Cells NuGet package to a .NET project and use it to merge several Excel files into one workbook.
// Use Cases: Create temporary worksheets and merge them into a master report. | Automate monthly spreadsheet consolidation for financial analysis. | Validate merge results by reading specific cells after combining files.
// AI Prompts: Generate C# code that installs Aspose.Cells via NuGet and merges a list of .xlsx files using CellsHelper.MergeFiles, including error handling and cache file cleanup. | Refactor the example to scan a directory for all Excel files and merge them into a single workbook. | Explain how to configure the cache file location and size when merging large numbers of worksheets with Aspose.Cells.

using System;
using Aspose.Cells;

// Step‑by‑step guide that shows how to add the Aspose.Cells library to a .NET project using the dotnet CLI, create sample workbooks, and combine them into a single file with CellsHelper.MergeFiles. The example loads the merged workbook, prints worksheet count, and displays A1 values for verification.
class Program
{
    static void Main()
    {
        // Add Aspose.Cells to the project via NuGet before compiling:
        //   dotnet add package Aspose.Cells

        // Prepare source Excel files (for demonstration they are created here)
        string[] sourceFiles = { "File1.xlsx", "File2.xlsx" };
        CreateSampleWorkbook(sourceFiles[0], "File 1 Content");
        CreateSampleWorkbook(sourceFiles[1], "File 2 Content");

        // Temporary cache file required by CellsHelper.MergeFiles
        string cacheFile = "Cache.tmp";

        // Destination file that will contain the merged result
        string destFile = "MergedOutput.xlsx";

        // Merge the source files into the destination file
        // Uses the CellsHelper.MergeFiles method (provided rule)
        CellsHelper.MergeFiles(sourceFiles, cacheFile, destFile);

        // Load the merged workbook to verify the result
        Workbook mergedWorkbook = new Workbook(destFile);
        Console.WriteLine("Merged workbook contains " + mergedWorkbook.Worksheets.Count + " worksheets.");
        Console.WriteLine("Sheet1 A1: " + mergedWorkbook.Worksheets[0].Cells["A1"].StringValue);
        if (mergedWorkbook.Worksheets.Count > 1)
        {
            Console.WriteLine("Sheet2 A1: " + mergedWorkbook.Worksheets[1].Cells["A1"].StringValue);
        }
    }

    // Helper method to create a simple workbook with a single cell value
    static void CreateSampleWorkbook(string filePath, string cellValue)
    {
        Workbook wb = new Workbook();                     // Create a new workbook
        wb.Worksheets[0].Cells["A1"].PutValue(cellValue); // Put a value in A1
        wb.Save(filePath);                                // Save the workbook
    }
}
