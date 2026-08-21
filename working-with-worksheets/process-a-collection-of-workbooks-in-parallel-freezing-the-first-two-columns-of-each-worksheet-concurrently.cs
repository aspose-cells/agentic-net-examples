// Title: Freeze First Two Columns in All Worksheets of Multiple Excel Workbooks in Parallel with Aspose.Cells for .NET
// Description: Loads each workbook from a supplied list, iterates through every worksheet, applies FreezePanes to lock columns A and B (rows remain unfrozen), saves the modified file to a target folder, and runs the operation concurrently using Parallel.ForEach. Errors are logged per file without stopping the batch.
// Keywords: Aspose.Cells | .NET | C# | FreezePanes | freeze columns | parallel processing | batch Excel | multiple workbooks | save to output folder | thread‑safe Excel automation
// Common Searches: Aspose.Cells freeze first two columns | How to freeze columns in all sheets using C# | Parallel processing Excel files Aspose.Cells | Batch freeze panes Aspose.Cells .NET | Save processed workbooks to another directory C#
// Developer Intent: Programmatically apply a column freeze to the first two columns of every worksheet across a collection of Excel files simultaneously and write the updated files to a separate location.
// Use Cases: Batch‑process a folder of generated reports so identifier columns stay visible while scrolling. | Integrate into a data‑export pipeline that must freeze columns in thousands of workbooks without blocking the main thread. | Create read‑only copies of existing workbooks with frozen columns for distribution while preserving the originals. | Automate preparation of template workbooks for multiple users, ensuring consistent column locking across all sheets.
// AI Prompts: Generate C# code that uses Aspose.Cells to freeze the first two columns of every sheet in a list of Excel files and process them concurrently. | Show how to add robust logging and exception handling when processing Excel workbooks in parallel with Aspose.Cells. | Demonstrate customizing the output path while preserving original filenames for batch‑processed workbooks.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// Loads each workbook from a supplied list, iterates through every worksheet, applies FreezePanes to lock columns A and B (rows remain unfrozen), saves the modified file to a target folder, and runs the operation concurrently using Parallel.ForEach. Errors are logged per file without stopping the batch.
public class WorkbookProcessor
{
    // Processes a collection of Excel files in parallel,
    // freezing the first two columns of every worksheet in each file.
    public static void ProcessWorkbooks(IEnumerable<string> inputFiles, string outputDirectory)
    {
        // Ensure the output directory exists
        Directory.CreateDirectory(outputDirectory);

        // Parallel processing of each workbook file
        Parallel.ForEach(inputFiles, inputFile =>
        {
            try
            {
                // Verify the input file exists before loading
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"File not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Freeze the first two columns (A and B)
                    // row = 0 (no rows frozen), column = 2 (C column), totalRows = 0, totalColumns = 2
                    sheet.FreezePanes(0, 2, 0, 2);
                }

                // Determine the output file path (same name, different folder)
                string outputPath = Path.Combine(outputDirectory, Path.GetFileName(inputFile));

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any errors for this file without stopping other tasks
                Console.WriteLine($"Error processing file '{inputFile}': {ex.Message}");
            }
        });
    }

    // Example usage
    public static void Main()
    {
        // List of Excel files to process
        List<string> files = new List<string>
        {
            @"C:\Data\Book1.xlsx",
            @"C:\Data\Book2.xlsx",
            @"C:\Data\Book3.xlsx"
        };

        // Folder where processed files will be saved
        string outputFolder = @"C:\Data\Processed";

        try
        {
            ProcessWorkbooks(files, outputFolder);
            Console.WriteLine("All workbooks have been processed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
