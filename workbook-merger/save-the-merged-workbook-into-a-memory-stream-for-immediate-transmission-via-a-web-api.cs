// Title: Combine two Excel workbooks into a single XLSX MemoryStream with Aspose.Cells in C# for web API responses
// AI Prompts: Write a C# method that loads two .xlsx files, merges them using Workbook.Combine, and returns the merged workbook as a MemoryStream ready for an HTTP response. | Generate code that checks the existence of two workbook paths, combines the workbooks with Aspose.Cells, saves the result to a MemoryStream in Xlsx format, and resets the stream position. | Create a reusable function that accepts two Excel file paths, merges the workbooks without creating intermediate files, and outputs the combined workbook as a MemoryStream.
// Common Searches: asp.net core return merged excel workbook as memory stream using aspose.cells | c# combine two xlsx files into a memory stream for API response | how to use Aspose.Cells Combine method and save to MemoryStream | merge excel workbooks without writing to disk in C#
// Tags: Aspose.Cells workbook combine to memory stream | C# merge Excel workbooks Xlsx stream | save combined workbook as Xlsx MemoryStream | web API return Excel memory stream Aspose.Cells | validate workbook file paths Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads two Excel workbooks, merges them with Aspose.Cells' Combine API, saves the merged workbook to a MemoryStream in XLSX format, resets the stream position, and returns it for immediate transmission via a web API.
public class WorkbookMerger
{
    // Merges two workbooks and returns the combined workbook as a MemoryStream (XLSX format).
    public MemoryStream MergeWorkbooksToStream(string firstFilePath, string secondFilePath)
    {
        try
        {
            // Verify that both input files exist.
            if (!File.Exists(firstFilePath))
                throw new FileNotFoundException($"File not found: {firstFilePath}");
            if (!File.Exists(secondFilePath))
                throw new FileNotFoundException($"File not found: {secondFilePath}");

            // Load the destination workbook (the one that will receive the other workbook).
            Workbook destWorkbook = new Workbook(firstFilePath);

            // Load the source workbook (the one to be merged into the destination).
            Workbook sourceWorkbook = new Workbook(secondFilePath);

            // Combine the source workbook into the destination workbook.
            destWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook to a memory stream in XLSX format.
            MemoryStream resultStream = new MemoryStream();
            destWorkbook.Save(resultStream, SaveFormat.Xlsx);
            resultStream.Position = 0; // Reset position for downstream consumers.

            return resultStream;
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to provide context while preserving the original exception.
            throw new ApplicationException("An error occurred while merging workbooks.", ex);
        }
    }
}

// Example console usage.
public class Program
{
    public static void Main()
    {
        // Example file paths; replace with actual locations as needed.
        string filePath1 = "File1.xlsx";
        string filePath2 = "File2.xlsx";

        WorkbookMerger merger = new WorkbookMerger();

        try
        {
            using (MemoryStream mergedStream = merger.MergeWorkbooksToStream(filePath1, filePath2))
            {
                // Save the merged workbook to a file for demonstration purposes.
                string outputPath = "MergedWorkbook.xlsx";
                using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    mergedStream.CopyTo(file);
                }
                Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to merge workbooks: {ex.Message}");
        }
    }
}
