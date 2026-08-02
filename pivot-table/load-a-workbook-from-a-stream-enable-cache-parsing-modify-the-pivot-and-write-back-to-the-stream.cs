// Title: C# – Load Excel workbook from stream, enable pivot cache parsing, refresh all pivots, and write back using Aspose.Cells
// Description: Demonstrates how to use Aspose.Cells for .NET to load an Excel file from an input stream with LoadOptions.ParsingPivotCachedRecords enabled, refresh every pivot table via Worksheets.RefreshPivotTables(), and save the modified workbook directly to an output stream. Includes error handling and a sample console program.
// Keywords: Aspose.Cells | .NET | C# | LoadOptions | ParsingPivotCachedRecords | pivot cache | refresh pivot tables | stream processing | Workbook.SaveToStream | Excel automation
// Common Searches: Aspose.Cells load workbook from stream with pivot cache | Enable ParsingPivotCachedRecords in Aspose.Cells | Refresh all pivot tables programmatically .NET | Save Excel workbook to stream using Aspose.Cells | How to process Excel pivot tables in a web API | C# example for pivot cache parsing and refresh
// Developer Intent: Load an Excel workbook from a stream, turn on pivot cache parsing, refresh its pivot tables, and output the updated file to another stream.
// Use Cases: Web API that receives an Excel file stream, updates pivot data, and returns the refreshed file. | Scheduled service that reads Excel reports from a shared folder, refreshes embedded pivots, and stores the updated files. | Document conversion pipeline that streams Excel input, applies pivot refresh, and streams the result to downstream processors. | Desktop utility that batch‑processes multiple workbooks, refreshing pivots without loading entire files into memory.
// AI Prompts: Write C# code using Aspose.Cells to read an Excel file from a MemoryStream, enable ParsingPivotCachedRecords, refresh all pivot tables, and return the result as a byte array. | Show how to handle large Excel workbooks with pivot caches by streaming input and output in Aspose.Cells, including error handling. | Explain best practices for refreshing pivot tables after loading a workbook with pivot cache parsing enabled in Aspose.Cells for .NET. | Generate a console application example that accepts input and output file paths, processes pivot tables via streams, and logs exceptions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to use Aspose.Cells for .NET to load an Excel file from an input stream with LoadOptions.ParsingPivotCachedRecords enabled, refresh every pivot table via Worksheets.RefreshPivotTables(), and save the modified workbook directly to an output stream. Includes error handling and a sample console program.
public class PivotCacheProcessor
{
    // Loads a workbook from the input stream, refreshes pivot tables, and writes to the output stream.
    public static void Process(Stream inputStream, Stream outputStream)
    {
        try
        {
            // Enable parsing of pivot cached records.
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingPivotCachedRecords = true
            };

            // Load workbook with the specified options.
            Workbook workbook = new Workbook(inputStream, loadOptions);

            // Refresh all pivot tables in the workbook.
            workbook.Worksheets.RefreshPivotTables();

            // Save workbook to a memory stream and copy to the output stream.
            using (MemoryStream tempStream = workbook.SaveToStream())
            {
                tempStream.Position = 0;
                tempStream.CopyTo(outputStream);
                outputStream.Position = 0;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    // Entry point for the console application.
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: PivotCacheProcessor <inputFilePath> <outputFilePath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Verify input file exists.
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Open streams and process the workbook.
            using (FileStream inputStream = File.OpenRead(inputPath))
            using (FileStream outputStream = File.Create(outputPath))
            {
                PivotCacheProcessor.Process(inputStream, outputStream);
            }

            Console.WriteLine($"Processing completed. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
