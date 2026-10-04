// Title: Log the size of a merged Excel workbook after saving with Aspose.Cells for .NET
// AI Prompts: Write C# code that merges two .xlsx workbooks using Aspose.Cells, saves the combined workbook, and prints the saved file size in bytes. | Show how to retrieve and log the file size of a workbook after calling Workbook.Save in a .NET console application with Aspose.Cells.
// Common Searches: aspnet get file size after saving merged workbook with Aspose.Cells | c# log merged Excel file size after combining worksheets using Aspose.Cells | how to measure output size of merged workbook in Aspose.Cells .NET | retrieve byte length of saved workbook after Aspose.Cells Save method | monitor storage impact when merging multiple Excel files with Aspose.Cells
// Tags: Aspose.Cells merge workbooks and log file size | C# retrieve workbook byte length after Save | measure Excel output size with Aspose.Cells | log merged workbook storage impact .NET | Aspose.Cells SaveFormat.Xlsx file size monitoring

using System;
using System.IO;
using Aspose.Cells;

// // Merges two Excel workbooks using Aspose.Cells, saves the result as 'MergedWorkbook.xlsx', and writes the resulting file size in bytes to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook file paths
            string path1 = "Workbook1.xlsx";
            string path2 = "Workbook2.xlsx";

            // Verify that the source files exist
            if (!File.Exists(path1))
            {
                Console.WriteLine($"File not found: {path1}");
                return;
            }
            if (!File.Exists(path2))
            {
                Console.WriteLine($"File not found: {path2}");
                return;
            }

            // Load the workbooks to be merged
            Workbook wb1 = new Workbook(path1);
            Workbook wb2 = new Workbook(path2);

            // Create a new workbook that will contain the merged sheets
            Workbook merged = new Workbook();

            // Remove the default empty worksheet created with a new workbook
            merged.Worksheets.Clear();

            // Copy all worksheets from the first workbook into the merged workbook
            foreach (Worksheet ws in wb1.Worksheets)
            {
                merged.Worksheets.AddCopy(ws.Name);
            }

            // Copy all worksheets from the second workbook into the merged workbook
            foreach (Worksheet ws in wb2.Worksheets)
            {
                merged.Worksheets.AddCopy(ws.Name);
            }

            // Define the output file path and save the merged workbook
            string outputPath = "MergedWorkbook.xlsx";
            merged.Save(outputPath, SaveFormat.Xlsx);

            // After saving, obtain and log the file size to monitor storage impact
            FileInfo fileInfo = new FileInfo(outputPath);
            Console.WriteLine($"Merged workbook size: {fileInfo.Length} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
