// Title: Aspose.Cells for .NET – Convert Excel to CSV and split into 10,000‑row parts
// Description: C# sample that loads an Excel workbook with Aspose.Cells, determines the populated rows of the first worksheet, and writes a series of CSV files each limited to 10,000 rows. Files are named with part number and row range and saved to a user‑specified output folder.
// Keywords: Aspose.Cells | .NET | C# | Excel to CSV | split worksheet | CSV chunking | 10,000 rows | large Excel export | batch CSV generation | GitHub example
// Common Searches: Aspose.Cells split worksheet into multiple CSV files | C# export Excel rows to CSV in 10k batches | How to break a large Excel sheet into smaller CSV parts using .NET | Convert Excel to CSV with row limit using Aspose.Cells | Chunk large worksheet to CSV files programmatically
// Developer Intent: Generate separate CSV files, each containing up to 10,000 rows from the first worksheet of an Excel workbook.
// Use Cases: Export a massive sales report into manageable CSV chunks for downstream analytics. | Prepare data for databases that impose a maximum row count per import file. | Create parallel‑processing CSV partitions for a data‑pipeline or ETL workflow.
// AI Prompts: Write a C# method using Aspose.Cells that splits any worksheet into CSV files with a configurable row limit. | Show how to modify the example to process every worksheet in the workbook instead of only the first one. | Add a progress bar or detailed logging to the CSV splitting loop for long‑running jobs.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# sample that loads an Excel workbook with Aspose.Cells, determines the populated rows of the first worksheet, and writes a series of CSV files each limited to 10,000 rows. Files are named with part number and row range and saved to a user‑specified output folder.
    public class WorkbookToCsvSplitter
    {
        // Splits a large worksheet into multiple CSV files, each containing up to 10,000 rows.
        public static void Run(string sourceFilePath, string outputFolder)
        {
            try
            {
                // Verify source file exists
                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"Source file not found: {sourceFilePath}");
                    return;
                }

                // Ensure output folder exists
                Directory.CreateDirectory(outputFolder);

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourceFilePath);

                // Work with the first worksheet
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Determine the total number of rows that contain data
                // MaxDataRow returns the last row index with data (zero‑based)
                int totalRows = sourceSheet.Cells.MaxDataRow + 1;

                const int rowsPerFile = 10000;
                int fileCount = (totalRows + rowsPerFile - 1) / rowsPerFile;

                for (int i = 0; i < fileCount; i++)
                {
                    int startRow = i * rowsPerFile;
                    int rowsInChunk = Math.Min(rowsPerFile, totalRows - startRow);

                    // Create a copy of the source workbook for this chunk
                    Workbook chunkWorkbook = new Workbook();
                    sourceWorkbook.Copy(chunkWorkbook);
                    Worksheet chunkSheet = chunkWorkbook.Worksheets[0];

                    // Remove rows after the desired chunk
                    int rowsAfter = totalRows - (startRow + rowsInChunk);
                    if (rowsAfter > 0)
                    {
                        chunkSheet.Cells.DeleteRows(startRow + rowsInChunk, rowsAfter);
                    }

                    // Remove rows before the desired chunk
                    if (startRow > 0)
                    {
                        chunkSheet.Cells.DeleteRows(0, startRow);
                    }

                    // Build the output CSV file name
                    string csvFileName = Path.Combine(
                        outputFolder,
                        $"Part_{i + 1}_Rows_{startRow + 1}_to_{startRow + rowsInChunk}.csv");

                    // Save the chunk as CSV
                    chunkWorkbook.Save(csvFileName, SaveFormat.Csv);
                    Console.WriteLine($"Saved: {csvFileName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during splitting: {ex.Message}");
            }
        }
    }

    // Simple entry point for testing the splitter
    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsExamples <sourceFilePath> <outputFolder>");
                return;
            }

            string sourceFilePath = args[0];
            string outputFolder = args[1];

            WorkbookToCsvSplitter.Run(sourceFilePath, outputFolder);
        }
    }
}
