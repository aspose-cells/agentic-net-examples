using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsDemo
{
    public class PivotCacheProcessor
    {
        /// <summary>
        /// Loads a workbook from the given input stream with pivot cache parsing enabled,
        /// refreshes the first pivot table, and returns the modified workbook as a MemoryStream.
        /// </summary>
        /// <param name="inputStream">Stream containing the source Excel file.</param>
        /// <returns>MemoryStream containing the modified Excel file.</returns>
        public MemoryStream ProcessPivotCache(Stream inputStream)
        {
            try
            {
                // Ensure the input stream is at the beginning
                if (inputStream.CanSeek)
                    inputStream.Position = 0;

                // Enable parsing of pivot cached records
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    ParsingPivotCachedRecords = true
                };

                // Load the workbook using the options
                Workbook workbook = new Workbook(inputStream, loadOptions);

                // Access the first worksheet (assumed to contain a pivot table)
                Worksheet worksheet = workbook.Worksheets[0];

                // If there is at least one pivot table, modify it
                if (worksheet.PivotTables.Count > 0)
                {
                    PivotTable pivotTable = worksheet.PivotTables[0];

                    // Disable automatic updates (manual update)
                    pivotTable.ManualUpdate = true;

                    // Refresh the pivot data and recalculate
                    pivotTable.RefreshData();
                    pivotTable.CalculateData();
                }

                // Refresh all pivot tables in the workbook to ensure consistency
                workbook.Worksheets.RefreshPivotTables();

                // Save the modified workbook to a memory stream (default format)
                MemoryStream outputStream = workbook.SaveToStream();

                // Reset the position of the output stream for downstream consumption
                if (outputStream.CanSeek)
                    outputStream.Position = 0;

                return outputStream;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing pivot cache: {ex.Message}");
                throw;
            }
        }
    }

    internal class Program
    {
        /// <summary>
        /// Entry point. Expects two arguments: input file path and output file path.
        /// </summary>
        private static void Main(string[] args)
        {
            try
            {
                if (args.Length != 2)
                {
                    Console.WriteLine("Usage: AsposeCellsDemo <input.xlsx> <output.xlsx>");
                    return;
                }

                string inputPath = args[0];
                string outputPath = args[1];

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Process the workbook
                using (FileStream inputStream = File.OpenRead(inputPath))
                {
                    PivotCacheProcessor processor = new PivotCacheProcessor();
                    MemoryStream resultStream = processor.ProcessPivotCache(inputStream);

                    // Write the result to the output file
                    using (FileStream outputStream = File.Create(outputPath))
                    {
                        resultStream.CopyTo(outputStream);
                    }

                    Console.WriteLine($"Processed workbook saved to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}