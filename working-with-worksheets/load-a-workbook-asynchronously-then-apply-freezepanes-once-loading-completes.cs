// Title: Asynchronously load an Excel workbook, apply FreezePanes to the first row and column, and save it with Aspose.Cells for .NET
// AI Prompts: Read an Excel file into a MemoryStream using async I/O, load it with Aspose.Cells Workbook, freeze the first row and column on the first worksheet, and save the result without blocking the calling thread. | After the workbook is loaded asynchronously, change the FreezePanes parameters to lock rows 5‑10 and columns B‑E on the target worksheet. | Add logic that creates the output directory if it does not exist before invoking Workbook.Save inside an async method. | Wrap the Workbook.Save call in Task.Run, capture any exceptions, and return a Task that signals completion of the async processing.
// Common Searches: asp.net core async load excel file with aspose.cells and apply freeze panes | how to freeze first row and column after loading workbook asynchronously in c# | non‑blocking save of modified excel workbook using aspose.cells | read excel into memory stream asynchronously then set FreezePanes Aspose.Cells | c# async workbook processing Aspose.Cells freeze panes example
// Tags: asynchronous Excel workbook loading Aspose.Cells | FreezePanes method Aspose.Cells worksheet | memory stream workbook load .NET | non‑blocking workbook save Aspose.Cells | error handling file not found Aspose.Cells | async file I/O with Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// This example demonstrates how to read an Excel file into a MemoryStream using asynchronous I/O, load it with Aspose.Cells, apply FreezePanes to lock the first row and column on the first worksheet, and save the modified workbook without blocking the calling thread.
public class WorkbookProcessor
{
    // Asynchronously loads a workbook, applies FreezePanes, and saves the result.
    public static async Task ProcessWorkbookAsync(string inputPath, string outputPath)
    {
        try
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Read the source file into a memory stream asynchronously.
            using (var memoryStream = new MemoryStream())
            {
                // Open the source file for asynchronous reading.
                using (FileStream sourceStream = new FileStream(
                    inputPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read,
                    bufferSize: 4096,
                    useAsync: true))
                {
                    await sourceStream.CopyToAsync(memoryStream).ConfigureAwait(false);
                }

                // Reset the position of the memory stream before loading.
                memoryStream.Position = 0;

                // Load the workbook from the memory stream.
                var workbook = new Workbook(memoryStream);

                // Apply FreezePanes to the first worksheet.
                // Freeze the first row and first column.
                Worksheet sheet = workbook.Worksheets[0];
                // totalRows, totalColumns, frozenRows, frozenColumns
                sheet.FreezePanes(1, 1, 1, 1);

                // Save the modified workbook to the specified output path.
                // Wrap synchronous Save in Task.Run to avoid blocking the calling thread.
                await Task.Run(() => workbook.Save(outputPath)).ConfigureAwait(false);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
            throw;
        }
    }

    // Example usage.
    public static async Task Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: WorkbookProcessor <inputPath> <outputPath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            await ProcessWorkbookAsync(inputPath, outputPath);
            Console.WriteLine("Workbook processed and saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to process workbook: {ex.Message}");
        }
    }
}
