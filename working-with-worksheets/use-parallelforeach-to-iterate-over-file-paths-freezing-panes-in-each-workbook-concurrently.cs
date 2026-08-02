// Title: Parallel Freeze Panes in Multiple Excel Workbooks with Aspose.Cells (C#)
// Description: Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, freezes panes at cell C3 on the first worksheet, saves the file, and processes all workbooks concurrently using Parallel.ForEach.
// Keywords: Aspose.Cells | C# | Parallel.ForEach | freeze panes | batch Excel processing | multithreaded workbook update | concurrent Excel manipulation | worksheet view settings
// Common Searches: freeze panes in all Excel files using C# | parallel processing of Excel workbooks Aspose.Cells | batch freeze panes at C3 | how to use Parallel.ForEach with Aspose.Cells | set freeze panes for multiple workbooks concurrently
// Developer Intent: Apply a freeze‑pane setting to the first worksheet of every .xlsx file in a directory in a single, parallel operation.
// Use Cases: Prepare a set of financial reports so header rows stay visible when users scroll. | Automate view configuration for thousands of data exports before publishing to a web portal. | Accelerate a CI/CD pipeline that updates workbook display options across generated files.
// AI Prompts: Generate C# code that uses Aspose.Cells to freeze panes at a configurable cell for all .xlsx files in a folder, employing Parallel.ForEach with robust error handling and logging. | Show how to implement thread‑safe file access and exception aggregation when processing many workbooks in parallel with Aspose.Cells. | Modify the example to freeze panes on a worksheet identified by name rather than the first sheet while keeping the parallel processing logic.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// Scans a folder for .xlsx files, loads each workbook with Aspose.Cells, freezes panes at cell C3 on the first worksheet, saves the file, and processes all workbooks concurrently using Parallel.ForEach.
class FreezePanesParallel
{
    static void Main()
    {
        // Define the folder containing Excel files
        string inputFolder = @"C:\InputFolder";

        // Verify that the folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Get all .xlsx files in the folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        if (excelFiles.Length == 0)
        {
            Console.WriteLine("No Excel files found in the specified folder.");
            return;
        }

        // Process each workbook in parallel
        Parallel.ForEach(excelFiles, filePath =>
        {
            try
            {
                // Ensure the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Freeze panes at cell C3 (row index 2, column index 2)
                worksheet.FreezePanes(2, 2, 2, 2);

                // Save changes back to the same file
                workbook.Save(filePath);

                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        });

        Console.WriteLine("All workbooks have been processed.");
    }
}
