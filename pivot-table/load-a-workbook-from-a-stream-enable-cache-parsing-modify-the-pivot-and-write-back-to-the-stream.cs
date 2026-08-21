// Title: Load Excel workbook from a stream, enable pivot cache parsing, refresh pivots, and save to MemoryStream using Aspose.Cells for .NET
// Description: Demonstrates how to load an XLSX workbook from a Stream with LoadOptions.ParsingPivotCachedRecords enabled, refresh all pivot tables, adjust the first pivot's ManualUpdate setting, recalculate its data, and write the updated workbook back to a MemoryStream for further processing.
// Keywords: Aspose.Cells load from stream | ParsingPivotCachedRecords | refresh pivot tables C# | modify pivot ManualUpdate | save workbook to MemoryStream | Aspose.Cells pivot cache | C# Excel pivot processing
// Common Searches: Aspose.Cells enable pivot cache parsing when loading workbook | C# refresh all pivot tables programmatically | How to set ManualUpdate false for a pivot table using Aspose.Cells | Save modified Excel file to MemoryStream in .NET | Process Excel stream and update pivots Aspose.Cells
// Developer Intent: Load an Excel file from a stream, turn on pivot cache parsing, update and recalculate pivot tables, and return the modified workbook as a MemoryStream.
// Use Cases: Web API endpoint that receives an uploaded XLSX, refreshes its pivots, and returns the updated file as a byte array. | Automated nightly job that opens stored workbooks, refreshes pivot data, and stores the result back to a database BLOB. | Cloud function that reads an Excel stream from storage, updates pivot tables, and uploads the revised workbook without creating temporary files.
// AI Prompts: Write C# code that loads an Excel workbook from a Stream with Aspose.Cells, enables ParsingPivotCachedRecords, refreshes all pivot tables, sets the first pivot's ManualUpdate to false, recalculates it, and returns a MemoryStream. | Provide a robust try‑catch example that ensures a MemoryStream is always returned when processing pivot tables with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to load an XLSX workbook from a Stream with LoadOptions.ParsingPivotCachedRecords enabled, refresh all pivot tables, adjust the first pivot's ManualUpdate setting, recalculate its data, and write the updated workbook back to a MemoryStream for further processing.
public class PivotCacheProcessor
{
    // Processes a workbook stream: enables pivot cache parsing, refreshes/updates pivots,
    // and returns the modified workbook as a memory stream.
    public static void Process(Stream inputStream, out MemoryStream outputStream)
    {
        try
        {
            // Enable parsing of pivot cached records.
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                ParsingPivotCachedRecords = true
            };

            // Load the workbook from the provided stream.
            var workbook = new Workbook(inputStream, loadOptions);

            // Refresh all pivot tables to reflect any source data changes.
            workbook.Worksheets.RefreshPivotTables();

            // Example of additional pivot modification.
            if (workbook.Worksheets[0].PivotTables.Count > 0)
            {
                var pivot = workbook.Worksheets[0].PivotTables[0];
                pivot.ManualUpdate = false;                     // Enable automatic updates.
                pivot.RefreshData();                            // Refresh the pivot's cached data.
                pivot.CalculateData();                         // Recalculate the pivot results.
            }

            // Save the modified workbook to a memory stream.
            outputStream = workbook.SaveToStream();
        }
        catch
        {
            // Ensure outputStream is always assigned.
            outputStream = new MemoryStream();
            throw;
        }
    }
}

public class Program
{
    // Entry point required for the console application.
    public static void Main(string[] args)
    {
        try
        {
            // Define input and output file paths.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Process the workbook.
            using (FileStream inputStream = File.OpenRead(inputPath))
            {
                PivotCacheProcessor.Process(inputStream, out MemoryStream resultStream);

                // Write the result to the output file.
                resultStream.Position = 0;
                using (FileStream outputStream = File.Create(outputPath))
                {
                    resultStream.CopyTo(outputStream);
                }
            }

            Console.WriteLine($"Workbook processed successfully. Output saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
