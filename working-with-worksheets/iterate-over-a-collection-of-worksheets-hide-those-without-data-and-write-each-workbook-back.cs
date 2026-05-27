using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetVisibilityProcessor
{
    class Program
    {
        static void Main()
        {
            // Folder containing the workbooks to process
            string inputFolder = @"C:\InputWorkbooks";
            // Folder where the processed workbooks will be saved
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the input folder
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string filePath in workbookFiles)
            {
                // Load the workbook (create rule)
                Workbook workbook = new Workbook(filePath);

                // Iterate over each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine if the worksheet contains any data
                    // MaxDataRow and MaxDataColumn are -1 when the sheet is empty
                    bool hasData = sheet.Cells.MaxDataRow >= 0 && sheet.Cells.MaxDataColumn >= 0;

                    // Hide the worksheet if it has no data
                    if (!hasData)
                    {
                        // Use the IsVisible property to hide the sheet
                        sheet.IsVisible = false;
                    }
                }

                // Build the output file path (same name, different folder)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the modified workbook (save rule)
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }

            Console.WriteLine("Processing completed.");
        }
    }
}