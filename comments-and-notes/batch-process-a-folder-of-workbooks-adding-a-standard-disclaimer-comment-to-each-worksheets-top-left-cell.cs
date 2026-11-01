// Title: Add a confidentiality disclaimer comment to cell A1 of every worksheet in multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: how to programmatically add the same comment to A1 in all sheets of several Excel files with Aspose.Cells C# | batch process a folder of .xlsx and .xls files to insert a disclaimer note in the top‑left cell using Aspose.Cells | C# Aspose.Cells add comment to each worksheet's first cell while overwriting the original workbook | iterate through worksheets and add a comment with Aspose.Cells .NET for multiple workbooks | automate adding a confidentiality comment to Excel workbooks in a directory using Aspose.Cells
// Tags: add comment to A1 Aspose.Cells .NET | batch process Excel workbooks C# | insert disclaimer note worksheet comment | overwrite original workbook Aspose.Cells | iterate worksheets add comment C#

using System;
using System.IO;
using Aspose.Cells;

// The C# console app receives a folder path, scans it for .xlsx and .xls files, loads each workbook with Aspose.Cells, iterates through every worksheet, adds a comment containing a confidentiality disclaimer to cell A1, saves the workbook back to the same file, and logs processing results or errors.
class BatchDisclaimerAdder
{
    // Standard disclaimer text to be added as a comment.
    private const string DisclaimerText = "This workbook is confidential and intended for authorized personnel only.";

    static void Main(string[] args)
    {
        // Validate input folder path.
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the folder containing workbooks.");
            return;
        }

        string folderPath = args[0];

        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Process each Excel file in the folder (including .xlsx and .xls).
        string[] excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls")
                continue; // Skip non‑Excel files.

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Add a comment to the top‑left cell (A1).
                    int commentIndex = sheet.Comments.Add("A1");
                    Comment comment = sheet.Comments[commentIndex];
                    comment.Author = "System";
                    comment.Note = DisclaimerText;
                }

                // Save the workbook, overwriting the original file.
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
