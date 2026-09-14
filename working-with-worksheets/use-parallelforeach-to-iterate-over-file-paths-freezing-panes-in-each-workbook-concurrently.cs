// Title: Use Parallel.ForEach to freeze the top row and first column in multiple Excel workbooks with Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a folder for .xlsx files, processes each workbook concurrently with Parallel.ForEach, applies Worksheet.FreezePanes to lock the first row and column, and saves the changes. | Adapt the sample to add robust error handling for missing files and I/O exceptions while performing concurrent FreezePanes operations using Aspose.Cells.
// Common Searches: c# parallel processing freeze panes aspose.cells multiple workbooks | how to apply freeze panes to first row and column in batch of Excel files using Aspose.Cells | concurrent processing of .xlsx files with Aspose.Cells FreezePanes method | best practice for thread‑safe workbook updates with Aspose.Cells in .NET
// Tags: Aspose.Cells FreezePanes parallel execution | batch freeze panes .xlsx C# | thread‑safe workbook modification Aspose.Cells | exception handling for concurrent Excel processing | enumerate .xlsx files directory C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The example enumerates all .xlsx files in a folder, then uses Parallel.ForEach to open each workbook with Aspose.Cells, applies Worksheet.FreezePanes(0,0,1,1) to lock the first row and column of the first worksheet, saves the workbook back to the same file, and logs successes or errors while safely handling missing files and other exceptions.
class Program
{
    static void Main()
    {
        // Retrieve workbook file paths to process
        IEnumerable<string> filePaths = GetWorkbookFilePaths();

        // Process each workbook in parallel with safety checks
        Parallel.ForEach(filePaths, path =>
        {
            try
            {
                // Ensure the file exists before attempting to load
                if (!File.Exists(path))
                {
                    Console.WriteLine($"File not found: {path}");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(path);

                // Access the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze the top row and leftmost column (cell A2)
                // FreezePanes(totalRows, totalColumns, rows, columns)
                // Using 0 for totalRows/totalColumns lets Aspose calculate automatically.
                sheet.FreezePanes(0, 0, 1, 1);

                // Save the workbook back to the same file (overwrites original)
                workbook.Save(path);
                Console.WriteLine($"Processed: {path}");
            }
            catch (Exception ex)
            {
                // Log any errors for this file without stopping other tasks
                Console.WriteLine($"Error processing '{path}': {ex.Message}");
            }
        });
    }

    // Placeholder method to retrieve workbook file paths.
    // Replace with actual implementation as needed.
    static IEnumerable<string> GetWorkbookFilePaths()
    {
        // Example: return all .xlsx files in a directory
        string folder = @"C:\Workbooks";
        if (!Directory.Exists(folder))
        {
            Console.WriteLine($"Directory not found: {folder}");
            return new List<string>();
        }

        return Directory.EnumerateFiles(folder, "*.xlsx", SearchOption.TopDirectoryOnly);
    }
}
