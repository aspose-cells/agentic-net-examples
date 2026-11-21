// Title: Batch hide blank worksheets in Excel workbooks and overwrite the files with Aspose.Cells for .NET
// Common Searches: C# code to hide empty sheets in multiple Excel files using Aspose.Cells | how to detect and hide blank worksheets with MaxDataRow in Aspose.Cells | process a list of workbook paths and save changes in place with Aspose.Cells .NET | batch hide worksheets without data and overwrite original XLSX files in C# | Aspose.Cells hide worksheet visibility based on cell content programmatically
// Tags: hide blank worksheets Aspose.Cells | detect empty sheet MaxDataRow C# | batch hide empty sheets Aspose.Cells | save workbook in place Aspose.Cells | set worksheet visibility programmatically .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookProcessor
{
    // The WorksheetHider class iterates through a collection of Excel file paths, loads each workbook with Aspose.Cells, checks each worksheet for data using Cells.MaxDataRow, hides worksheets that are empty by setting IsVisible to false, and saves the modified workbook back to the original XLSX file while handling missing files and runtime errors.
    public class WorksheetHider
    {
        /// <param name="workbookPaths">Collection of full file paths to the workbooks.</param>
        public void HideEmptyWorksheets(IEnumerable<string> workbookPaths)
        {
            foreach (var path in workbookPaths)
            {
                if (!File.Exists(path))
                {
                    Console.WriteLine($"File not found: {path}");
                    continue;
                }

                try
                {
                    // Load the workbook from the specified file.
                    Workbook workbook = new Workbook(path);

                    // Iterate through all worksheets in the workbook.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Determine if the worksheet contains any data.
                        // MaxDataRow returns -1 when there are no cells with data.
                        bool isEmpty = sheet.Cells.MaxDataRow == -1;

                        // Hide the worksheet if it is empty.
                        sheet.IsVisible = !isEmpty;
                    }

                    // Save the modified workbook back to the original file.
                    workbook.Save(path, SaveFormat.Xlsx);
                    Console.WriteLine($"Processed workbook: {path}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{path}': {ex.Message}");
                }
            }
        }
    }

    public class Program
    {
        // Entry point for the console application.
        public static void Main(string[] args)
        {
            // If no arguments provided, inform the user.
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Please provide at least one workbook file path as an argument.");
                return;
            }

            var hider = new WorksheetHider();
            hider.HideEmptyWorksheets(args);
        }
    }
}
