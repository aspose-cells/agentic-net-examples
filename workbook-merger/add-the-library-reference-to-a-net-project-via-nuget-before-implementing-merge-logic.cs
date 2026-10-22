// Title: Install Aspose.Cells via NuGet and merge two Excel workbooks into one using a C# console application
// AI Prompts: Show the exact dotnet CLI commands to add the Aspose.Cells NuGet package to a .NET console project, then generate C# code that merges two .xlsx files into a single workbook. | Write a C# console snippet that loads two Excel files with Aspose.Cells, copies every worksheet into a new workbook using Worksheets.AddCopy, and saves the merged result as Merged.xlsx.
// Common Searches: how to add Aspose.Cells NuGet package and merge Excel files in a C# console app | C# example for merging multiple .xlsx workbooks with Aspose.Cells AddCopy | using Aspose.Cells to combine worksheets from two workbooks into one file
// Tags: install Aspose.Cells NuGet package | merge Excel workbooks Aspose.Cells C# | copy worksheets between workbooks Aspose.Cells | C# console workbook consolidation | Aspose.Cells workbook merging example

using System;
using System.IO;
using Aspose.Cells;

// The console application verifies the presence of two source .xlsx files, loads them with Aspose.Cells, creates a new workbook, copies all worksheets from both source workbooks using Worksheets.AddCopy, saves the merged workbook as Merged.xlsx, and handles missing files or runtime errors.
class WorkbookMerger
{
    static void Main(string[] args)
    {
        // Paths to the source Excel files
        string sourcePath1 = "Source1.xlsx";
        string sourcePath2 = "Source2.xlsx";

        // Path for the merged output file
        string mergedPath = "Merged.xlsx";

        try
        {
            // Verify that source files exist
            if (!File.Exists(sourcePath1))
            {
                Console.WriteLine($"Error: File not found - {sourcePath1}");
                return;
            }

            if (!File.Exists(sourcePath2))
            {
                Console.WriteLine($"Error: File not found - {sourcePath2}");
                return;
            }

            // Load the source workbooks
            Workbook wb1 = new Workbook(sourcePath1);
            Workbook wb2 = new Workbook(sourcePath2);

            // Create a new workbook to hold the merged content
            Workbook mergedWorkbook = new Workbook();

            // Remove the default empty worksheet created by Aspose.Cells
            mergedWorkbook.Worksheets.Clear();

            // Copy all worksheets from the first workbook into the merged workbook
            foreach (Worksheet ws in wb1.Worksheets)
            {
                // AddCopy expects the source worksheet name
                mergedWorkbook.Worksheets.AddCopy(ws.Name);
            }

            // Copy all worksheets from the second workbook into the merged workbook
            foreach (Worksheet ws in wb2.Worksheets)
            {
                mergedWorkbook.Worksheets.AddCopy(ws.Name);
            }

            // Save the merged workbook to the specified file
            mergedWorkbook.Save(mergedPath);
            Console.WriteLine($"Workbooks merged successfully into '{mergedPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
