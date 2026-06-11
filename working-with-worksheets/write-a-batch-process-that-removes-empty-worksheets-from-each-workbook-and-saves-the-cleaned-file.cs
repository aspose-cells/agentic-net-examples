using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatch
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input folder and output folder
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsBatch <inputFolder> <outputFolder>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Process all Excel files in the input folder (including subfolders if needed)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                // Filter supported Excel formats
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                    continue;

                try
                {
                    // Load the workbook (uses the provided load rule)
                    Workbook workbook = new Workbook(filePath);

                    // Iterate worksheets in reverse order to safely remove them
                    for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                    {
                        Worksheet sheet = workbook.Worksheets[i];
                        // A worksheet is considered empty if it has no data rows and no data columns
                        bool isEmpty = sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1;
                        if (isEmpty)
                        {
                            // Remove the empty worksheet (uses the provided RemoveAt rule)
                            workbook.Worksheets.RemoveAt(i);
                        }
                    }

                    // Prepare output file path (same file name, different folder)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the cleaned workbook (uses the provided save rule)
                    workbook.Save(outputPath);
                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}