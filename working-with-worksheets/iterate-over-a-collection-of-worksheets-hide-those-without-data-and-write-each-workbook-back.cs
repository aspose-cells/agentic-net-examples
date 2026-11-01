// Title: Batch hide empty worksheets in multiple Excel workbooks and overwrite files using Aspose.Cells for .NET
// AI Prompts: Generate C# code that receives an IEnumerable of Excel file paths, loads each workbook with Aspose.Cells, evaluates each worksheet’s Cells.MaxDataRow and Cells.MaxDataColumn to determine emptiness, sets Worksheet.IsVisible to false for empty sheets, and saves the workbook back to its original location. | Create a reusable method in C# that iterates through all worksheets of a workbook, hides any sheet that contains no data, and overwrites the source .xlsx file using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# hide blank sheets in a batch of Excel files | programmatically set worksheet visibility based on MaxDataRow MaxDataColumn Aspose.Cells | overwrite original workbook after changing sheet visibility with Aspose.Cells .NET | how to iterate through worksheets and conceal empty ones using Aspose.Cells | C# process multiple .xlsx files to remove empty worksheets automatically
// Tags: conceal empty worksheets Aspose.Cells C# | batch process Excel workbooks Aspose.Cells | Worksheet.IsVisible based on Cells.MaxDataRow MaxDataColumn | overwrite original .xlsx file Aspose.Cells | detect blank sheet using Cells.MaxDataRow Cells.MaxDataColumn

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookProcessor
{
    // // Loads each workbook from a list of file paths, checks every worksheet for data using MaxDataRow/MaxDataColumn, hides worksheets that are empty, and saves the workbook back, overwriting the original file.
    public class WorksheetHider
    {
        /// <param name="workbookPaths">Full paths to the workbook files to process.</param>
        public void ProcessWorkbooks(IEnumerable<string> workbookPaths)
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
                    // Load the workbook from the file system
                    var workbook = new Workbook(path);

                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Determine if the worksheet contains any data
                        // MaxDataRow and MaxDataColumn return -1 when the sheet is empty
                        bool hasData = sheet.Cells.MaxDataRow >= 0 && sheet.Cells.MaxDataColumn >= 0;

                        // Hide the worksheet if it has no data
                        sheet.IsVisible = hasData;
                    }

                    // Save the modified workbook back to the original file (overwrite)
                    workbook.Save(path);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{path}': {ex.Message}");
                }
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            var processor = new WorksheetHider();

            // List of workbook file paths to process
            var workbooks = new List<string>
            {
                @"C:\Data\Report1.xlsx",
                @"C:\Data\Report2.xlsx",
                @"C:\Data\Report3.xlsx"
            };

            processor.ProcessWorkbooks(workbooks);
        }
    }
}
