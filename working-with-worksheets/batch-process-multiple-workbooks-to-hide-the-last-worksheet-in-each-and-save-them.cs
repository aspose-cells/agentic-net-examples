// Title: How to batch hide the last worksheet in multiple Excel workbooks and save them using Aspose.Cells for .NET
// Common Searches: C# code to hide the last sheet in a batch of .xlsx files with Aspose.Cells | process multiple Excel workbooks to set last worksheet visibility to hidden using Aspose.Cells .NET | save modified workbooks to a different folder after hiding worksheets in Aspose.Cells | automate hiding the last worksheet across several Excel files in C# | Aspose.Cells programmatic way to hide a worksheet in many workbooks
// Tags: batch hide worksheet Aspose.Cells .NET | set worksheet IsVisible false C# | process multiple Excel workbooks Aspose.Cells | save workbook to alternate directory Aspose.Cells | automate worksheet visibility Excel .NET

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to iterate over an array of Excel file paths, load each workbook with Aspose.Cells, hide the last worksheet by setting its IsVisible property to false, ensure the output directory exists, and save the updated workbook to a specified folder while preserving the original file name.
class WorkbookBatchProcessor
{
    /// <param name="inputFiles">Array of full paths to the source workbooks.</param>
    /// <param name="outputFolder">Folder where the processed workbooks will be saved.</param>
    public static void HideLastWorksheetAndSave(string[] inputFiles, string outputFolder)
    {
        // Ensure the output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        foreach (string inputPath in inputFiles)
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Get the collection of worksheets
            WorksheetCollection sheets = workbook.Worksheets;

            // Proceed only if there is at least one worksheet
            if (sheets.Count > 0)
            {
                // Index of the last worksheet (zero‑based)
                int lastIndex = sheets.Count - 1;

                // Hide the last worksheet
                sheets[lastIndex].IsVisible = false;
            }

            // Determine the output file path (same file name, different folder)
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }

    // Example usage
    static void Main()
    {
        // Define the source workbook files (could be retrieved from a directory, etc.)
        string[] sourceFiles = Directory.GetFiles(@"C:\InputWorkbooks", "*.xlsx");

        // Define where to save the processed workbooks
        string destinationFolder = @"C:\ProcessedWorkbooks";

        // Process the workbooks
        HideLastWorksheetAndSave(sourceFiles, destinationFolder);

        Console.WriteLine("Processing completed.");
    }
}
