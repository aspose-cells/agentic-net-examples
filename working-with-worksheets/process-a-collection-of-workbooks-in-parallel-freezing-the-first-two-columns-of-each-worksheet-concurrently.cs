// Title: Freeze columns A and B in every worksheet of multiple Excel files concurrently using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a directory for .xlsx files, opens each workbook with Aspose.Cells, applies FreezePanes to lock the first two columns on all worksheets, and saves the changes while executing the processing in parallel. | Create a parallel batch routine in C# that uses Aspose.Cells to load each Excel workbook from a folder, freeze columns A and B on every sheet, handle per‑file errors, and write the updated files to an output folder.
// Common Searches: how to apply FreezePanes to the first two columns of all sheets in a batch of Excel workbooks with Aspose.Cells C# | parallel processing of multiple .xlsx files to freeze columns using Aspose.Cells .NET | C# example for freezing columns A and B across worksheets in many workbooks | Aspose.Cells freeze first two columns in each worksheet while saving files in parallel
// Tags: freeze panes first two columns Aspose.Cells | batch freeze columns Excel C# Aspose.Cells | parallel workbook processing Aspose.Cells .NET | apply FreezePanes to all worksheets .xlsx | load and save multiple Excel files C# Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample enumerates all .xlsx files in a given folder, loads each workbook with Aspose.Cells, freezes columns A and B on every worksheet using the FreezePanes method, saves the modified workbook to an output directory, and performs the entire operation concurrently with Parallel.ForEach while handling errors per file.
class Program
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string inputFolder = @"C:\Workbooks\Input";
        // Folder where the processed workbooks will be saved (can be the same as inputFolder)
        string outputFolder = @"C:\Workbooks\Output";

        // Ensure the output folder exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the input folder
        string[] workbookPaths = Directory.GetFiles(inputFolder, "*.xlsx");

        // Process each workbook in parallel
        Parallel.ForEach(workbookPaths, workbookPath =>
        {
            try
            {
                // Verify the file exists before loading
                if (!File.Exists(workbookPath))
                {
                    Console.Error.WriteLine($"File not found: '{workbookPath}'");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Freeze the first two columns (column index is zero‑based, so 2 means columns A and B)
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // FreezePanes(row, column, totalRows, totalColumns)
                    sheet.FreezePanes(0, 2, 0, 2);
                }

                // Determine the output file path
                string fileName = Path.GetFileName(workbookPath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the modified workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Handle exceptions per workbook (logging, etc.)
                Console.Error.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
            }
        });

        Console.WriteLine("Processing completed.");
    }
}
