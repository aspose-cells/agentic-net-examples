using System;
using Aspose.Cells;

public class WorkbookCombiner
{
    // Combines multiple Excel files into a single workbook.
    public static void CombineWorkbooks(string[] sourceFiles, string outputFile)
    {
        // Create an empty destination workbook.
        Workbook destWorkbook = new Workbook();

        // Load each source workbook and merge it into the destination.
        foreach (string filePath in sourceFiles)
        {
            // Load the source workbook from file.
            Workbook srcWorkbook = new Workbook(filePath);

            // Merge the source workbook into the destination workbook.
            destWorkbook.Combine(srcWorkbook);
        }

        // Save the combined workbook to the specified output path.
        destWorkbook.Save(outputFile, SaveFormat.Xlsx);
    }

    // Example entry point demonstrating usage.
    public static void Main()
    {
        // List of workbook files to be combined.
        string[] filesToCombine = new string[]
        {
            "File1.xlsx",
            "File2.xlsx",
            "File3.xlsx"
        };

        // Path for the resulting combined workbook.
        string combinedFilePath = "CombinedWorkbook.xlsx";

        // Perform the combination.
        CombineWorkbooks(filesToCombine, combinedFilePath);

        Console.WriteLine($"Workbooks combined successfully. Output saved to: {combinedFilePath}");
    }
}