using System;
using System.IO;
using Aspose.Cells;

class BatchDisclaimerProcessor
{
    // Standard disclaimer text to be added as a comment
    private const string DisclaimerText = "Disclaimer: This workbook is confidential and intended for authorized personnel only.";

    static void Main()
    {
        // Folder containing the Excel workbooks to process
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where the processed workbooks will be saved (can be the same as inputFolder to overwrite)
        string outputFolder = @"C:\ProcessedWorkbooks";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each Excel file in the input folder (supports .xls, .xlsx, .xlsm)
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue; // Skip non‑Excel files

            // Load the workbook using the provided constructor (load rule)
            Workbook workbook = new Workbook(filePath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add a comment to cell A1 (top‑left cell) if it doesn't already exist
                int commentIndex = sheet.Comments.Add("A1");
                sheet.Comments[commentIndex].Note = DisclaimerText;
                sheet.Comments[commentIndex].Author = Environment.UserName;
            }

            // Determine the output file path
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

            // Save the modified workbook using the provided Save method (save rule)
            workbook.Save(outputPath);
        }

        Console.WriteLine("Batch processing completed.");
    }
}