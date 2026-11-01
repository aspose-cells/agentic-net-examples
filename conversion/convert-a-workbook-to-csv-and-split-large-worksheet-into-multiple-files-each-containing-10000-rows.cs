// Title: Convert an Excel workbook to CSV and split a large worksheet into 10,000‑row CSV files using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells split large Excel sheet into multiple CSV files each containing 10,000 rows | how to export an Excel worksheet to CSV in 10k‑row chunks with Aspose.Cells for .NET | Aspose.Cells copy rows and preserve column widths when saving a worksheet as CSV | generate several CSV parts from one workbook using Aspose.Cells C# example | convert Excel to CSV and break data into 10,000‑row files with Aspose.Cells
// Tags: Aspose.Cells copy rows to new workbook C# | export worksheet to CSV preserving column widths Aspose.Cells | split Excel sheet into chunked CSV files .NET | save workbook as CSV format using Aspose.Cells | process large Excel data in 10k row batches C# | Aspose.Cells CSV generation for large worksheets

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input.xlsx workbook, determines the populated row and column range of the first worksheet, and iterates through the data in 10,000‑row blocks. For each block it creates a new workbook, copies rows together with their styles and column widths, and saves the chunk as a CSV file (output_part_1.csv, output_part_2.csv, …). It logs progress and includes basic file‑existence and exception handling.
class WorkbookCsvSplitter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Determine the total number of rows and columns that contain data
            int totalRows = sourceSheet.Cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
            int totalColumns = sourceSheet.Cells.MaxDataColumn + 1; // MaxDataColumn is zero‑based

            const int rowsPerFile = 10000;
            int partIndex = 1;

            // Process rows in chunks of 10,000
            for (int startRow = 0; startRow < totalRows; startRow += rowsPerFile)
            {
                int rowsToCopy = Math.Min(rowsPerFile, totalRows - startRow);

                // Create a new workbook for the current chunk
                Workbook chunkWorkbook = new Workbook();
                Worksheet chunkSheet = chunkWorkbook.Worksheets[0];

                // Copy rows from source sheet to the chunk sheet
                // The overload copies rows along with their styles
                chunkSheet.Cells.CopyRows(sourceSheet.Cells, startRow, 0, rowsToCopy);

                // Copy column widths to preserve layout
                for (int col = 0; col < totalColumns; col++)
                {
                    double width = sourceSheet.Cells.GetColumnWidth(col);
                    chunkSheet.Cells.SetColumnWidth(col, width);
                }

                // Save the chunk as a CSV file
                string outputFileName = $"output_part_{partIndex}.csv";
                chunkWorkbook.Save(outputFileName, SaveFormat.Csv);

                Console.WriteLine($"Created {outputFileName} with rows {startRow + 1} to {startRow + rowsToCopy}");
                partIndex++;
            }

            Console.WriteLine("Splitting completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
