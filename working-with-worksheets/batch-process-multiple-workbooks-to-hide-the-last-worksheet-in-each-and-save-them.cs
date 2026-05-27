using System;
using System.IO;
using Aspose.Cells;

namespace BatchHideLastWorksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source workbooks
            string sourceFolder = @"C:\InputWorkbooks";
            // Folder where the modified workbooks will be saved
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Load the workbook (uses the Workbook(string) constructor rule)
                Workbook workbook = new Workbook(filePath);

                // Get the index of the last worksheet
                int lastIndex = workbook.Worksheets.Count - 1;
                if (lastIndex >= 0)
                {
                    // Hide the last worksheet (uses Worksheet.SetVisible method rule)
                    workbook.Worksheets[lastIndex].SetVisible(false, true);
                }

                // Build the output file path
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the modified workbook (uses the Workbook.Save(string) rule)
                workbook.Save(outputPath);
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}