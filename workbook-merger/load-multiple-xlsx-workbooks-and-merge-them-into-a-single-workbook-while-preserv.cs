using System;
using Aspose.Cells;

public class WorkbookMerger
{
    // Merges multiple XLSX files into a single workbook.
    public static void MergeWorkbooks(string[] sourceFiles, string destinationFile)
    {
        // Create an empty workbook that will receive the combined data.
        Workbook mergedWorkbook = new Workbook();

        foreach (string filePath in sourceFiles)
        {
            // Load each source workbook from disk.
            Workbook sourceWorkbook = new Workbook(filePath);

            // Combine the source workbook into the merged workbook.
            mergedWorkbook.Combine(sourceWorkbook);
        }

        // Save the merged workbook to the specified destination path in XLSX format.
        mergedWorkbook.Save(destinationFile, SaveFormat.Xlsx);
    }

    // Example entry point demonstrating usage.
    public static void Main()
    {
        string[] filesToMerge = { "File1.xlsx", "File2.xlsx", "File3.xlsx" };
        string outputPath = "CombinedWorkbook.xlsx";

        try
        {
            MergeWorkbooks(filesToMerge, outputPath);
            Console.WriteLine($"Workbooks merged successfully into '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during merging: {ex.Message}");
        }
    }
}