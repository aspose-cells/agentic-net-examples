// Title: Freeze the first two columns in every worksheet of multiple Excel workbooks concurrently with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a collection of .xlsx files, uses Aspose.Cells to freeze the first two columns on each worksheet, and saves the workbooks while processing them in parallel with Parallel.ForEach. | Adapt the example to freeze the first three rows instead of columns, preserving the parallel processing structure. | Add comprehensive error handling that logs any file that fails to load or save and ensures the parallel loop continues with remaining workbooks. | Create a reusable method that accepts a column count parameter and applies FreezePanes to all worksheets in a batch of Excel files processed concurrently.
// Common Searches: how to use Aspose.Cells to freeze columns in all sheets of multiple Excel files in C# | parallel batch processing of .xlsx workbooks with FreezePanes using Aspose.Cells | C# example for freezing first two columns across worksheets in a folder of Excel workbooks | Aspose.Cells concurrent workbook modification freeze panes performance | process many Excel files simultaneously and apply column freeze with .NET
// Tags: Aspose.Cells FreezePanes column freeze batch processing | parallel Excel workbook modification C# Aspose.Cells | freeze first two columns all worksheets Aspose.Cells | concurrent processing of .xlsx files with Aspose.Cells | batch apply FreezePanes to multiple workbooks .NET | error handling in parallel Aspose.Cells file loop

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// Processes a set of .xlsx files in parallel using Parallel.ForEach, loads each workbook with Aspose.Cells, freezes the first two columns of every worksheet via FreezePanes(0,2,0,0), and saves the modified files to a designated output folder.
public class WorkbookProcessor
{
    /// <param name="inputFiles">Full paths of the source Excel files.</param>
    /// <param name="outputFolder">Folder where the processed files will be saved.</param>
    public void ProcessWorkbooksParallel(IEnumerable<string> inputFiles, string outputFolder)
    {
        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Parallel processing of each workbook
        Parallel.ForEach(inputFiles, inputFile =>
        {
            try
            {
                // Skip missing files
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"File not found: {inputFile}");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Freeze the first two columns (row index 0, column index 2)
                    // Aspose.Cells overload requires four parameters; set totalRows and totalColumns to 0
                    sheet.FreezePanes(0, 2, 0, 0);
                }

                // Build the output file path (preserve original file name)
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputFile));

                // Save the modified workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{inputFile}': {ex.Message}");
            }
        });
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Define input and output directories
            string inputFolder = @"C:\InputExcels";
            string outputFolder = @"C:\ProcessedExcels";

            // Verify input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input directory does not exist: {inputFolder}");
                return;
            }

            // Get all .xlsx files from the input directory
            IEnumerable<string> files = Directory.GetFiles(inputFolder, "*.xlsx");

            // Process the workbooks
            WorkbookProcessor processor = new WorkbookProcessor();
            processor.ProcessWorkbooksParallel(files, outputFolder);

            Console.WriteLine("Workbook processing completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
