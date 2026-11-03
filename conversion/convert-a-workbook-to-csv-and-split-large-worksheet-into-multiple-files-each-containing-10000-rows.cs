// Title: Split a large Excel worksheet into 10,000‑row CSV files using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# split worksheet into multiple CSV files with 10k rows each | How to export Excel data to CSV in 10,000‑row chunks using Aspose.Cells .NET | C# code to convert a large XLSX workbook to several CSV files of 10000 rows | Saving parts of an Excel sheet as separate CSV files with Aspose.Cells | Chunked CSV export from Excel workbook using TxtSaveOptions in C#
// Tags: Aspose.Cells split worksheet to CSV chunks | Csv export with TxtSaveOptions in C# | Chunked CSV generation from Excel using Aspose.Cells | Copy cells between workbooks Aspose.Cells | Large Excel to multiple CSV files .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsSplitCsv
{
    // The example loads an XLSX workbook, determines the total rows in the first worksheet, and iterates through the data in 10,000‑row blocks. For each block it creates a new workbook, copies non‑null cells (preserving values and styles), and saves the block as a CSV file using TxtSaveOptions with SaveFormat.CSV. Output files are named sequentially (output_part_1.csv, output_part_2.csv, …). The code includes file existence checks and exception handling for robust execution.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                Workbook sourceWorkbook;

                // Attempt to load the workbook; handle password‑protected files gracefully
                try
                {
                    sourceWorkbook = new Workbook(inputPath);
                }
                catch (CellsException ex)
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }

                // Use the first worksheet as the source
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Total rows that contain data (zero‑based, so add 1 for count)
                int totalRows = sourceSheet.Cells.MaxDataRow + 1;

                const int rowsPerFile = 10000; // Maximum rows per CSV file
                int fileIndex = 1;

                // Process rows in chunks
                for (int startRow = 0; startRow < totalRows; startRow += rowsPerFile)
                {
                    int endRow = Math.Min(startRow + rowsPerFile - 1, totalRows - 1);

                    // Create a new workbook for the current chunk
                    Workbook chunkWorkbook = new Workbook();
                    Worksheet chunkSheet = chunkWorkbook.Worksheets[0];

                    // Copy cells from source to chunk workbook
                    for (int r = startRow; r <= endRow; r++)
                    {
                        for (int c = 0; c <= sourceSheet.Cells.MaxDataColumn; c++)
                        {
                            Cell srcCell = sourceSheet.Cells[r, c];
                            if (srcCell != null && srcCell.Type != CellValueType.IsNull)
                            {
                                Cell destCell = chunkSheet.Cells[r - startRow, c];
                                destCell.PutValue(srcCell.Value);
                                // Preserve style (optional)
                                destCell.SetStyle(srcCell.GetStyle());
                            }
                        }
                    }

                    // Save the chunk as CSV
                    TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
                    string outputFileName = $"output_part_{fileIndex}.csv";
                    chunkWorkbook.Save(outputFileName, csvOptions);
                    Console.WriteLine($"Saved: {outputFileName}");

                    fileIndex++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
