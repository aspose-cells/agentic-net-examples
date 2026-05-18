using System;
using System.IO;
using Aspose.Cells;

namespace MergeWorkbooksDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths of the source Excel files to be merged
                string[] sourceFiles = new string[]
                {
                    "Source1.xlsx",
                    "Source2.xlsx",
                    "Source3.xlsx"
                };

                // Create the destination workbook (empty workbook)
                Workbook destinationWorkbook = new Workbook();

                // Load each source workbook using the Workbook(string) constructor
                foreach (string filePath in sourceFiles)
                {
                    // Verify that the source file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Warning: Source file not found and will be skipped: {filePath}");
                        continue;
                    }

                    // Load source workbook
                    Workbook sourceWorkbook = new Workbook(filePath);

                    // Combine the source workbook into the destination workbook
                    destinationWorkbook.Combine(sourceWorkbook);
                }

                // Save the merged workbook to a new file
                string outputPath = "MergedWorkbook.xlsx";
                destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbooks merged successfully. Output saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during merging: {ex.Message}");
            }
        }
    }
}