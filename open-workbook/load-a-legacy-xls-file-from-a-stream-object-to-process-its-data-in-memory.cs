// Title: Load an Excel 97‑2003 (.xls) workbook from a Stream with Aspose.Cells for .NET and read its cells in memory
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions to open an .xls file from a Stream and prints every used cell value. | Show how to catch and log exceptions when loading a legacy XLS workbook from a Stream with Aspose.Cells. | Provide an example that iterates over the used range of the first worksheet after loading an Excel 97‑2003 file from a Stream.
// Common Searches: aspnet load xls from memory stream using aspose.cells | c# read legacy Excel 97-2003 workbook without saving to disk | how to use LoadOptions Excel97To2003 with Aspose.Cells | iterate over used cells in an XLS workbook loaded from a stream | exception handling for loading XLS stream Aspose.Cells .NET
// Tags: load xls stream Aspose.Cells | LoadOptions Excel97To2003 | iterate used cells Aspose.Cells worksheet | in‑memory processing of legacy XLS | exception handling Aspose.Cells load

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates loading a legacy Excel 97‑2003 (.xls) file directly from a Stream using Aspose.Cells LoadOptions, accessing the first worksheet, determining its used range, and printing each cell's value to the console while handling potential errors.
public class ExcelProcessor
{
    // Loads a legacy XLS file from a Stream and processes its data in memory
    public void ProcessXlsFromStream(Stream xlsStream)
    {
        try
        {
            // Set load options to explicitly treat the stream as an Excel 97-2003 file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

            // Load the workbook from the provided stream using the load options
            Workbook workbook = new Workbook(xlsStream, loadOptions);

            // Access the first worksheet (you can change the index as needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Example processing: iterate through all used cells and output their values
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    object value = cells[row, col].Value;
                    Console.Write($"{value}\t");
                }
                Console.WriteLine();
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing XLS stream: {ex.Message}");
        }
    }
}

public class Program
{
    // Entry point for the application
    public static void Main(string[] args)
    {
        // Expect a file path as the first argument; fallback to a default name
        string filePath = args.Length > 0 ? args[0] : "sample.xls";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.Error.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Open the file as a read‑only stream and process it
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                ExcelProcessor processor = new ExcelProcessor();
                processor.ProcessXlsFromStream(fs);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
