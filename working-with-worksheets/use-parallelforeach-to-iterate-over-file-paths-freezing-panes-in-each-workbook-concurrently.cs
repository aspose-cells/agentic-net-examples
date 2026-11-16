// Title: Freeze the first row of each worksheet in multiple .xlsx files concurrently with Parallel.ForEach and Aspose.Cells for .NET
// Common Searches: how to apply FreezePanes to many Excel workbooks in parallel using Aspose.Cells C# | parallel processing of .xlsx files to freeze top row with Parallel.ForEach | batch freeze panes in Excel files using Aspose.Cells and multithreading | C# code to freeze first row in all workbooks in a folder concurrently | using Aspose.Cells to freeze panes while iterating files with Parallel.ForEach
// Tags: parallel FreezePanes with Aspose.Cells | batch freeze top row in .xlsx files | Aspose.Cells concurrent workbook modification | C# Parallel.ForEach Excel automation | freeze first row worksheet Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The example scans a folder for *.xlsx files, then uses Parallel.ForEach to load each workbook with Aspose.Cells, freezes the first row of the first worksheet via FreezePanes(0,0,1,0), saves the workbook back to the same file, and logs success or errors for each file.
class Program
{
    static void Main()
    {
        // Directory containing the Excel files
        string folderPath = @"C:\ExcelFiles";

        // Verify that the directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Get all Excel file paths (adjust the search pattern as needed)
        string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

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

                // Load the workbook from the current file
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze the top row (rowsToFreeze = 1, columnsToFreeze = 0)
                // totalRows and totalColumns can be set to 0 to use the default sheet size
                sheet.FreezePanes(0, 0, 1, 0);

                // Save the changes back to the same file
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        });

        Console.WriteLine("All workbooks have been processed.");
    }
}
