using System;
using System.IO;
using Aspose.Cells;

namespace BatchDisclaimer
{
    class Program
    {
        // Standard disclaimer text to be added as a comment.
        private const string DisclaimerText = "Confidential: This workbook is for internal use only.";

        static void Main(string[] args)
        {
            // Folder containing the Excel files to process.
            string inputFolder = @"C:\InputWorkbooks";

            // Optional: folder to save the modified workbooks.
            // If you want to overwrite the original files, set outputFolder = inputFolder.
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Process all .xlsx files in the input folder (non‑recursive).
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Load the workbook using the provided constructor rule.
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Add a comment to cell A1 (top‑left cell) with the disclaimer.
                    int commentIndex = sheet.Comments.Add("A1");
                    sheet.Comments[commentIndex].Note = DisclaimerText;
                    sheet.Comments[commentIndex].Author = "System";
                }

                // Determine the output file path (same name, different folder or overwrite).
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the modified workbook using the provided Save method.
                workbook.Save(outputPath);

                // Release resources.
                workbook.Dispose();
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}