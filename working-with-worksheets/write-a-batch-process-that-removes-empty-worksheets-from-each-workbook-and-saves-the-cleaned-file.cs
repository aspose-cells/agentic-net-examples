// Title: C# Batch Tool to Delete Empty Worksheets from Excel Files with Aspose.Cells
// Description: A C# console utility that scans a folder for Excel workbooks, loads each file with Aspose.Cells, removes every worksheet that has no data (MaxDataRow = -1), and saves the cleaned workbook to a target directory while preserving the original file name and format.
// Keywords: Aspose.Cells | C# | remove empty worksheets | batch delete blank sheets | Excel workbook cleanup | process folder of workbooks | save cleaned Excel files | .NET Excel automation | MaxDataRow | supported extensions xlsx xls xlsm xlsb csv
// Common Searches: C# batch remove blank worksheets Aspose.Cells | delete empty sheets from multiple Excel files .NET | how to clean Excel workbooks programmatically | Aspose.Cells remove worksheets with no data | automate Excel folder processing C#
// Developer Intent: Iterate through all Excel files in a directory, strip out worksheets that contain no data, and write the cleaned workbooks to an output folder.
// Use Cases: Archive a reports repository by eliminating unnecessary blank tabs to reduce file size. | Pre‑process user‑uploaded Excel files before importing them into a data pipeline, ensuring no empty sheets cause validation errors. | Create a scheduled maintenance job that keeps shared Excel libraries tidy by automatically removing empty worksheets.
// AI Prompts: Generate a version of EmptyWorksheetCleaner that logs each removed worksheet name to a CSV log file. | Adapt the batch process to retain at least one sheet when all are empty, adding a placeholder sheet called "Summary" with a custom message. | Write NUnit tests for EmptyWorksheetCleaner that confirm empty sheets are deleted and sheets with data remain unchanged.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatch
{
    // A C# console utility that scans a folder for Excel workbooks, loads each file with Aspose.Cells, removes every worksheet that has no data (MaxDataRow = -1), and saves the cleaned workbook to a target directory while preserving the original file name and format.
    public static class EmptyWorksheetCleaner
    {
        /// <param name="inputFolder">Folder containing source workbooks.</param>
        /// <param name="outputFolder">Folder where cleaned workbooks will be saved.</param>
        public static void ProcessFolder(string inputFolder, string outputFolder)
        {
            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Supported Excel extensions (add more if needed)
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".csv" };

            // Enumerate files with supported extensions
            foreach (string filePath in Directory.GetFiles(inputFolder))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue; // Skip non‑Excel files

                // Load the workbook (uses the provided Workbook(string) constructor)
                Workbook workbook = new Workbook(filePath);

                // Iterate worksheets in reverse order to safely remove items
                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // A worksheet is considered empty if it has no data rows
                    // MaxDataRow returns -1 when there is no data
                    if (sheet.Cells.MaxDataRow == -1)
                    {
                        // Remove the empty worksheet (uses WorksheetCollection.RemoveAt(int))
                        workbook.Worksheets.RemoveAt(i);
                    }
                }

                // Build the output file path (preserve original file name)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the cleaned workbook (uses Workbook.Save(string))
                workbook.Save(outputPath);
            }
        }

        // Example entry point
        public static void Main(string[] args)
        {
            // Example usage:
            // args[0] = input folder, args[1] = output folder
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: EmptyWorksheetCleaner <inputFolder> <outputFolder>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];

            ProcessFolder(inputFolder, outputFolder);
            Console.WriteLine("Processing completed.");
        }
    }
}
