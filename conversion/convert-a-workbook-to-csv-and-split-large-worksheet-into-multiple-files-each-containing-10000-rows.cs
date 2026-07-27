// Title: Aspose.Cells .NET – Split an Excel worksheet into 10,000‑row CSV files
// Description: C# example that loads an Excel workbook with Aspose.Cells, reads the first worksheet, and writes the data to a series of CSV files. Each file contains up to 10,000 rows, is named sequentially (Part_001.csv, Part_002.csv, …), and is saved to a user‑specified folder. The code handles missing files, creates the output directory, and logs progress.
// Keywords: Aspose.Cells CSV export | split Excel to multiple CSV | C# workbook to CSV chunks | large worksheet to CSV parts | 10k rows per CSV Aspose | Excel to CSV batch conversion | Aspose.Cells .NET example
// Common Searches: how to split Excel sheet into several CSV files using Aspose.Cells | Aspose.Cells export first worksheet to CSV in 10,000 row batches | C# convert large workbook to multiple CSV files | Aspose.Cells split worksheet by row count | save Excel rows as separate CSV files .NET
// Developer Intent: The developer needs to convert the first worksheet of an Excel workbook into a set of CSV files, each limited to 10,000 rows, using Aspose.Cells for .NET.
// Use Cases: Break a massive data sheet into smaller CSV files for tools that impose a row‑limit per upload. | Create chunked CSV outputs for parallel processing in ETL pipelines while preserving row order. | Generate manageable CSV parts for downstream analytics platforms that struggle with very large files.
// AI Prompts: Write a reusable C# method that splits any worksheet into CSV files with a configurable row limit using Aspose.Cells. | Show how to extend the sample to process all worksheets in a workbook and include the header row in each CSV part. | Suggest best‑practice logging and exception handling for missing source files, empty worksheets, and permission errors.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that loads an Excel workbook with Aspose.Cells, reads the first worksheet, and writes the data to a series of CSV files. Each file contains up to 10,000 rows, is named sequentially (Part_001.csv, Part_002.csv, …), and is saved to a user‑specified folder. The code handles missing files, creates the output directory, and logs progress.
    public class WorkbookToCsvSplitter
    {
        // Splits the first worksheet of a workbook into multiple CSV files,
        // each containing up to 10,000 rows.
        public static void Run(string sourceFilePath, string outputFolder)
        {
            // Verify source file exists
            if (!File.Exists(sourceFilePath))
            {
                Console.WriteLine($"Source file not found: {sourceFilePath}");
                return;
            }

            // Ensure output folder exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            try
            {
                // Load the source workbook (lifecycle: load)
                using (Workbook sourceWorkbook = new Workbook(sourceFilePath))
                {
                    // Access the first worksheet (you can modify to handle multiple sheets)
                    Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                    Cells sourceCells = sourceSheet.Cells;

                    // Determine the last row that contains data in column A (index 0)
                    int lastDataRow = sourceCells.GetLastDataRow(0);
                    // If the worksheet is empty, nothing to do
                    if (lastDataRow < 0)
                    {
                        Console.WriteLine("The worksheet contains no data.");
                        return;
                    }

                    const int rowsPerFile = 10000;
                    int fileIndex = 1;
                    int startRow = 0;

                    // Loop through the rows in chunks of 10,000
                    while (startRow <= lastDataRow)
                    {
                        // Calculate how many rows to copy for this chunk
                        int rowsToCopy = Math.Min(rowsPerFile, lastDataRow - startRow + 1);

                        // Create a new workbook for the chunk (lifecycle: create)
                        using (Workbook chunkWorkbook = new Workbook())
                        {
                            // The workbook already contains one default worksheet
                            Worksheet chunkSheet = chunkWorkbook.Worksheets[0];
                            Cells chunkCells = chunkSheet.Cells;

                            // Copy the required rows from the source worksheet to the chunk worksheet
                            // CopyRows(sourceCells, sourceStartRow, destinationStartRow, rowCount)
                            chunkCells.CopyRows(sourceCells, startRow, 0, rowsToCopy);

                            // Build the output CSV file name
                            string outputFile = Path.Combine(
                                outputFolder,
                                $"Part_{fileIndex:D3}.csv");

                            // Save the chunk workbook as CSV (lifecycle: save)
                            chunkWorkbook.Save(outputFile, SaveFormat.Csv);

                            Console.WriteLine($"Saved rows {startRow} - {startRow + rowsToCopy - 1} to {outputFile}");
                        }

                        // Prepare for next iteration
                        startRow += rowsToCopy;
                        fileIndex++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during processing: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            // Expect two arguments: source file path and output folder
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <sourceFilePath> <outputFolder>");
                return;
            }

            string sourceFilePath = args[0];
            string outputFolder = args[1];

            try
            {
                WorkbookToCsvSplitter.Run(sourceFilePath, outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
