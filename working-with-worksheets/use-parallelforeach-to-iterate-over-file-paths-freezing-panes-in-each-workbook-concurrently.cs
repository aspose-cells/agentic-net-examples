// Title: Freeze the first row and column in multiple Excel workbooks concurrently using Parallel.ForEach and Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates a collection of .xlsx file paths, uses Parallel.ForEach to open each workbook with Aspose.Cells, applies FreezePanes(1,1,1,1) to every worksheet, and saves the file. | Adapt the sample to freeze panes only on the first worksheet of each workbook while still processing the files in parallel. | Add robust error handling that logs missing files and exceptions without breaking the Parallel.ForEach loop. | Show how to limit the degree of parallelism when freezing panes in a large batch of Excel workbooks.
// Common Searches: c# parallel foreach freeze panes aspose.cells multiple workbooks | how to apply FreezePanes to all worksheets in a batch using Aspose.Cells | using Parallel.ForEach to process Excel files with Aspose.Cells .NET | freeze top row and left column in each sheet of many Excel files programmatically | batch freeze panes aspose.cells performance parallel processing
// Tags: parallel freeze panes aspose.cells | batch freeze panes .net | apply FreezePanes to all worksheets aspose.cells | process multiple Excel workbooks concurrently c# | aspocells parallel workbook modification

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace FreezePanesParallel
{
    // The program loads a list of Excel file paths, processes each workbook in parallel with Parallel.ForEach, freezes the first row and column on every worksheet using FreezePanes(1,1,1,1), saves the workbook, and logs successes or errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Example list of Excel file paths to process
            List<string> excelFilePaths = new List<string>
            {
                @"C:\Data\Report1.xlsx",
                @"C:\Data\Report2.xlsx",
                @"C:\Data\Report3.xlsx"
                // Add more paths as needed
            };

            // Process each workbook concurrently
            Parallel.ForEach(excelFilePaths, filePath =>
            {
                // Ensure the file exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Freeze the first row and first column.
                        // Parameters: row, column, totalRows, totalColumns
                        // Setting row = 1, column = 1, totalRows = 1, totalColumns = 1
                        // freezes the top row and leftmost column.
                        sheet.FreezePanes(1, 1, 1, 1);
                    }

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath);
                    Console.WriteLine($"Processed and saved: {filePath}");
                }
                catch (Exception ex)
                {
                    // Log any errors that occur during processing
                    Console.WriteLine($"Error processing {filePath}: {ex.Message}");
                }
            });

            Console.WriteLine("All files have been processed.");
        }
    }
}
