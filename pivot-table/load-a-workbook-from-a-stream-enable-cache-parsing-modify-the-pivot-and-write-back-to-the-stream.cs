// Title: Load an XLSX workbook from a Stream with pivot cache parsing, modify a source cell, refresh all pivot tables, and save to a MemoryStream using Aspose.Cells for .NET
// AI Prompts: Generate a C# method that opens an Excel file from a Stream with ParsingPivotCachedRecords enabled, changes cell B2, calls RefreshPivotTables, and returns the workbook as a MemoryStream using Aspose.Cells. | Show how to enable pivot cache parsing when loading a workbook, modify a source data value, and write the updated file to a MemoryStream in C# with Aspose.Cells. | Create example code that accepts an input Stream, loads the workbook with pivot cache support, updates a pivot source cell, refreshes all pivots, and outputs a MemoryStream.
// Common Searches: Aspose.Cells enable ParsingPivotCachedRecords when opening an XLSX stream | C# update pivot source data and refresh pivots using Aspose.Cells | How to write an Aspose.Cells workbook to a MemoryStream after editing | Load Excel file from FileStream, change a cell, and refresh pivot tables in .NET | Impact of pivot cache parsing on performance in Aspose.Cells
// Tags: ParsingPivotCachedRecords option Aspose.Cells | update all pivot caches Aspose.Cells | modify pivot source cell B2 C# | save workbook to MemoryStream Aspose.Cells | load Excel workbook from stream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotCacheApp
{
    // Loads an XLSX workbook from a Stream with pivot cache parsing enabled, updates cell B2 in the source data, refreshes all pivot tables, and returns the modified workbook as a MemoryStream using Aspose.Cells for .NET.
    public class PivotCacheProcessor
    {
        // Loads a workbook from a stream, modifies source data, refreshes pivots, and returns a memory stream.
        public static void Process(Stream inputStream, out MemoryStream outputStream)
        {
            try
            {
                // Enable parsing of pivot cached records while loading.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    ParsingPivotCachedRecords = true
                };

                // Load workbook with the specified options.
                Workbook workbook = new Workbook(inputStream, loadOptions);

                // Example modification: change a cell that is part of the pivot source data.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["B2"].PutValue(999);

                // Refresh all pivot tables to reflect the change.
                workbook.Worksheets.RefreshPivotTables();

                // Save the modified workbook to a memory stream.
                outputStream = workbook.SaveToStream();
            }
            catch
            {
                outputStream = null;
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before attempting to load.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                using (FileStream inputStream = File.OpenRead(inputPath))
                {
                    PivotCacheProcessor.Process(inputStream, out MemoryStream outputStream);

                    if (outputStream != null)
                    {
                        // Write the resulting stream to the output file.
                        File.WriteAllBytes(outputPath, outputStream.ToArray());
                        Console.WriteLine($"Processed workbook saved to {outputPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
