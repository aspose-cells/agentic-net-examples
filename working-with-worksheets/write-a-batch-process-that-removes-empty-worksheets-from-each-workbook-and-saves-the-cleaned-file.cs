// Title: C# Batch Tool to Remove Empty Worksheets from Excel Workbooks with Aspose.Cells
// Description: A console application that scans a folder for .xlsx, .xls, and .xlsm files, loads each workbook using Aspose.Cells, deletes worksheets that contain no data (MaxDataRow = -1 and MaxDataColumn = -1), and saves the cleaned files to a target directory while preserving original names.
// Keywords: Aspose.Cells | C# remove empty worksheets | batch Excel cleanup | delete blank sheets | process multiple workbooks | WorksheetCollection.RemoveAt | Workbook.Save | Excel automation .NET | remove blank tabs | reduce Excel file size
// Common Searches: batch remove blank worksheets Aspose.Cells C# | delete empty sheets from all Excel files in a folder | C# script to clean multiple workbooks by removing empty tabs | Aspose.Cells remove worksheets with no data rows | how to automate Excel sheet cleanup with .NET
// Developer Intent: Build a command‑line utility that iterates through a directory of Excel files, strips out any worksheet that has no content, and writes the sanitized workbooks to an output folder.
// Use Cases: Prepare client‑submitted report bundles by stripping placeholder sheets before archiving. | Trim the size of automated Excel exports that include unnecessary blank tabs. | Integrate into a CI/CD pipeline to ensure only populated worksheets are packaged for deployment. | Maintain a clean template library by removing empty sheets from legacy files.
// AI Prompts: Add logging that records the names of all worksheets removed for each workbook. | Modify the program to guarantee at least one worksheet remains, creating a default "Summary" sheet when all are empty. | Extend the script to process subfolders recursively while preserving the original folder hierarchy in the output location.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcess
{
    // A console application that scans a folder for .xlsx, .xls, and .xlsm files, loads each workbook using Aspose.Cells, deletes worksheets that contain no data (MaxDataRow = -1 and MaxDataColumn = -1), and saves the cleaned files to a target directory while preserving original names.
    class RemoveEmptyWorksheets
    {
        static void Main(string[] args)
        {
            // Input folder containing workbooks to process
            string inputFolder = @"C:\InputWorkbooks";
            // Output folder where cleaned workbooks will be saved
            string outputFolder = @"C:\CleanedWorkbooks";

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the input folder (supports .xlsx and .xls)
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
                    continue; // Skip non‑Excel files

                // Load the workbook (uses Workbook(string) constructor)
                Workbook workbook = new Workbook(filePath);

                // Iterate worksheets in reverse order to safely remove items
                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    // A worksheet is considered empty when it has no data rows and no data columns
                    bool isEmpty = sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1;

                    if (isEmpty)
                    {
                        // Remove the empty worksheet (uses WorksheetCollection.RemoveAt(int))
                        workbook.Worksheets.RemoveAt(i);
                    }
                }

                // Determine output file path (overwrite original name in output folder)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the cleaned workbook (uses Workbook.Save(string))
                workbook.Save(outputPath);
                workbook.Dispose();

                Console.WriteLine($"Processed: {Path.GetFileName(filePath)} -> Saved cleaned file to {outputPath}");
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
