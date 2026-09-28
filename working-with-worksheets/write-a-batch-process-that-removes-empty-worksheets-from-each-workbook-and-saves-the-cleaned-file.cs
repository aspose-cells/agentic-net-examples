// Title: C# batch program to delete empty worksheets from all Excel workbooks in a folder using Aspose.Cells
// AI Prompts: Write a C# console application that scans a given directory for .xlsx files, loads each workbook with Aspose.Cells, removes worksheets whose MaxDataRow and MaxDataColumn are negative, adds a placeholder sheet if the workbook becomes empty, and saves the cleaned files to an output folder. | Update the removal logic so that worksheets containing only formatting (no cell values) are also treated as empty before the workbook is saved.
// Common Searches: how to programmatically delete blank sheets from multiple Excel files using Aspose.Cells in C# | batch remove empty worksheets from a folder of .xlsx files with Aspose.Cells .NET | C# code to clean Excel workbooks by removing sheets that have no data rows | Aspose.Cells remove worksheets with only formatting and no values | ensure at least one worksheet remains after deleting empty sheets in Aspose.Cells
// Tags: remove empty worksheets Aspose.Cells .NET | batch clean Excel workbooks C# | delete blank sheets programmatically Aspose.Cells | add placeholder worksheet if none remain Aspose.Cells | filter worksheets by MaxDataRow Aspose.Cells | process multiple .xlsx files Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookCleaner
{
    // A C# console utility that iterates over every .xlsx file in a source folder, uses Aspose.Cells to identify and delete worksheets with no data (or only formatting), guarantees at least one sheet remains, and writes the cleaned workbooks to a target directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to process
            string inputFolder = @"C:\InputWorkbooks";

            // Folder where cleaned workbooks will be saved
            string outputFolder = @"C:\CleanedWorkbooks";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Collect indexes of worksheets that are empty
                List<int> emptySheetIndexes = new List<int>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // A worksheet is considered empty when it has no data rows and no data columns
                    if (sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0)
                    {
                        emptySheetIndexes.Add(sheet.Index);
                    }
                }

                // Remove empty worksheets starting from the highest index to avoid shifting issues
                emptySheetIndexes.Sort();
                emptySheetIndexes.Reverse();
                foreach (int index in emptySheetIndexes)
                {
                    workbook.Worksheets.RemoveAt(index);
                }

                // Aspose.Cells requires at least one worksheet; add a blank one if all were removed
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add();
                }

                // Save the cleaned workbook to the output folder (overwrites if exists)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                workbook.Save(outputPath);
            }

            Console.WriteLine("Processing complete. Cleaned workbooks are saved in: " + outputFolder);
        }
    }
}
