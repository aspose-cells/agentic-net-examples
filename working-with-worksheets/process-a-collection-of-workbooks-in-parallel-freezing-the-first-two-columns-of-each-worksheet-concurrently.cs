// Title: Freeze the first two columns of every worksheet in multiple Excel workbooks in parallel using Aspose.Cells for .NET
// Common Searches: C# code to freeze first two columns in every sheet of multiple Excel files with Aspose.Cells | How to apply FreezePanes to all worksheets while processing workbooks concurrently | Parallel.ForEach example for batch freezing columns in Excel workbooks using Aspose.Cells | Batch update Excel files to freeze columns A and B across all worksheets in .NET
// Tags: freeze panes first two columns Aspose.Cells .NET | parallel workbook processing Aspose.Cells | batch freeze columns Excel Aspose.Cells | apply FreezePanes to all worksheets C# | concurrent Excel worksheet modification Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace WorkbookProcessing
{
    // The example demonstrates a WorkbookProcessor class that receives a collection of Excel file paths, loads each workbook with Aspose.Cells, freezes columns A and B on every worksheet using FreezePanes(0,2,0,0), saves the workbook, and runs the entire operation concurrently via Parallel.ForEach while handling missing files and exceptions.
    public class WorkbookProcessor
    {
        /// <param name="workbookPaths">Collection of full file paths to the workbooks.</param>
        public void FreezeFirstTwoColumns(IEnumerable<string> workbookPaths)
        {
            Parallel.ForEach(workbookPaths, path =>
            {
                try
                {
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"File not found: {path}");
                        return;
                    }

                    // Load the workbook from the specified file.
                    var workbook = new Workbook(path);

                    // Freeze the first two columns (A and B) in each worksheet.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // row = 0 (no rows frozen), column = 2 (freeze columns before index 2), totalRows = 0, totalColumns = 0.
                        sheet.FreezePanes(0, 2, 0, 0);
                    }

                    // Save the changes back to the original file (overwrites the existing workbook).
                    workbook.Save(path);
                    Console.WriteLine($"Processed: {path}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing {path}: {ex.Message}");
                }
            });
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: replace with actual file paths or command‑line arguments.
            var files = new List<string>
            {
                @"C:\Books\Book1.xlsx",
                @"C:\Books\Book2.xlsx"
            };

            var processor = new WorkbookProcessor();
            processor.FreezeFirstTwoColumns(files);
        }
    }
}
