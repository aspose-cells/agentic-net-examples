// Title: Merge several .xls workbooks into a single .xlsx file using Aspose.Cells for .NET (C#)
// AI Prompts: Load each .xls file with `new Workbook(filePath)` and merge it into a destination workbook using the `Combine` method in C#. | Create an empty workbook, loop through a list of source file paths, handle missing files, combine all sheets, and save the final workbook as an .xlsx with `Workbook.Save`.
// Common Searches: c# aspnet load multiple xls files and merge into one xlsx using aspose.cells | how to combine several .xls workbooks into a single workbook with Aspose.Cells in .NET | Aspose.Cells combine workbooks from file paths with error handling | merge excel .xls files to .xlsx programmatically using Aspose.Cells C# | sample code for merging multiple Excel workbooks and saving as xlsx with Aspose.Cells
// Tags: combine workbooks Aspose.Cells C# | load xls workbook from file path Aspose.Cells | save merged workbook as xlsx Aspose.Cells | exception handling merging Excel files .NET | iterate source files merge Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace MergeWorkbooksExample
{
    // The example loads each .xls workbook using `new Workbook(filePath)`, merges them into an empty destination workbook with `Combine`, handles missing files and exceptions, and saves the consolidated result as a .xlsx file.
    class Program
    {
        static void Main()
        {
            // Paths of the source XLS workbooks to be merged
            string[] sourceFiles = new string[]
            {
                "Source1.xls",
                "Source2.xls",
                "Source3.xls"
            };

            // Create an empty destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Load each source workbook, combine it into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}");
                        continue;
                    }

                    Workbook sourceWorkbook = new Workbook(filePath);
                    destinationWorkbook.Combine(sourceWorkbook);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            // Save the merged workbook to a new file
            try
            {
                string outputPath = "MergedResult.xlsx";
                destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Merged workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save merged workbook: {ex.Message}");
            }
        }
    }
}
