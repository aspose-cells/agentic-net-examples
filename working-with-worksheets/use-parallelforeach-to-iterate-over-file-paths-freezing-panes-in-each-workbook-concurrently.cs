// Title: Freeze panes in multiple Excel workbooks concurrently with Parallel.ForEach – Aspose.Cells for .NET
// Description: A C# example that creates an output folder, loads each .xls/.xlsx file from a collection, freezes panes at cell C3 on the first worksheet, saves the workbook, and runs the whole process in parallel using Parallel.ForEach for high‑throughput Excel automation.
// Keywords: Aspose.Cells | Parallel.ForEach | C# | .NET | freeze panes | batch Excel processing | concurrent workbook modification | multi‑threaded Excel | Excel automation | performance optimization
// Common Searches: freeze panes in many Excel files using Aspose.Cells | Parallel.ForEach batch workbook example C# | how to apply freeze panes to multiple workbooks concurrently | Aspose.Cells parallel processing tutorial | C# code to freeze rows and columns in a folder of Excel files
// Developer Intent: Apply a freeze pane at cell C3 to every workbook in a supplied list and save the results in parallel.
// Use Cases: Standardize header freezing across dozens of monthly report files with a single command. | Speed up a cloud service that receives bulk Excel uploads by freezing panes in each file concurrently before further analysis. | Migrate legacy spreadsheets where a uniform pane freeze is required, processing an entire directory in minutes instead of hours.
// AI Prompts: Write C# code that uses Aspose.Cells to freeze panes at a user‑defined cell for a collection of Excel files processed with Parallel.ForEach, including error handling and output folder creation. | Show how to modify the sample to target a different worksheet index or custom row/column coordinates while keeping parallel execution. | Give recommendations for safely scaling parallel workbook processing with Aspose.Cells, covering thread‑safety, memory usage, and optimal degree of parallelism.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Cells;

// A C# example that creates an output folder, loads each .xls/.xlsx file from a collection, freezes panes at cell C3 on the first worksheet, saves the workbook, and runs the whole process in parallel using Parallel.ForEach for high‑throughput Excel automation.
public class FreezePanesParallel
{
    // Processes a collection of Excel files in parallel, freezing panes in each workbook.
    public static void ProcessWorkbooks(IEnumerable<string> filePaths, string outputFolder)
    {
        // Ensure the output directory exists.
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        Parallel.ForEach(filePaths, filePath =>
        {
            try
            {
                // Verify the source file exists.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook from the given file path.
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns.
                worksheet.FreezePanes(2, 2, 2, 2);

                // Build the output file path.
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        });
    }
}

public class Program
{
    // Entry point of the application.
    public static void Main(string[] args)
    {
        try
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <inputFolder> <outputFolder>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];

            // Verify input folder exists.
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Collect all Excel files in the input folder.
            IEnumerable<string> excelFiles = Directory.EnumerateFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly)
                .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));

            // Process the workbooks.
            FreezePanesParallel.ProcessWorkbooks(excelFiles, outputFolder);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
