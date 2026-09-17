// Title: Convert an Excel workbook to CSV and split each worksheet into 10,000‑row files with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to load an XLSX workbook, iterate every worksheet, and create CSV files that contain no more than 10,000 rows each. | Demonstrate how to copy cell values and their styles from a source worksheet into a new workbook chunk and save that chunk as a CSV using Aspose.Cells SaveFormat.Csv. | Provide a loop that builds CSV filenames from the original workbook name, worksheet name, and part number while handling an arbitrary number of worksheets.
// Common Searches: Aspose.Cells C# split large worksheet into multiple CSV files with row limit | How to export Excel sheet to CSV in 10k‑row chunks using Aspose.Cells .NET | C# example for chunked CSV generation from an XLSX workbook with Aspose | Save each 10,000 rows of an Excel worksheet as separate CSV files using Aspose.Cells | Preserve cell formatting while converting Excel to CSV in chunks with Aspose.Cells
// Tags: Aspose.Cells CSV chunk export | C# Excel to CSV with row limit | 10k‑row CSV chunking using Aspose.Cells | preserve formatting in CSV export Aspose | batch CSV generation from large Excel

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an input.xlsx workbook, walks through every worksheet, calculates the total data rows, and processes the sheet in 10,000‑row blocks. For each block it creates a temporary workbook, copies cell values and styles, names the output file with the original workbook name, worksheet name, and part number, and saves the block as a CSV file using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook srcWorkbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet srcSheet in srcWorkbook.Worksheets)
        {
            // Determine the total number of rows that contain data
            int totalRows = srcSheet.Cells.MaxDataRow + 1; // MaxDataRow is zero‑based
            int chunkSize = 10000;                         // Rows per CSV file
            int partNumber = 1;

            // Process the worksheet in chunks of 10,000 rows
            for (int startRow = 0; startRow < totalRows; startRow += chunkSize, partNumber++)
            {
                // Create a new workbook that will hold the current chunk
                Workbook chunkWorkbook = new Workbook();
                Worksheet chunkSheet = chunkWorkbook.Worksheets[0];
                chunkSheet.Name = srcSheet.Name;

                // Number of rows to copy in this chunk
                int rowsInChunk = Math.Min(chunkSize, totalRows - startRow);

                // Copy cells from the source worksheet to the chunk worksheet
                for (int r = 0; r < rowsInChunk; r++)
                {
                    int srcRowIndex = startRow + r;
                    for (int c = 0; c <= srcSheet.Cells.MaxDataColumn; c++)
                    {
                        Cell srcCell = srcSheet.Cells[srcRowIndex, c];
                        if (srcCell != null && srcCell.Type != CellValueType.IsNull)
                        {
                            Cell destCell = chunkSheet.Cells[r, c];
                            destCell.PutValue(srcCell.Value);
                            // Preserve cell style (optional)
                            destCell.SetStyle(srcCell.GetStyle());
                        }
                    }
                }

                // Build the output CSV file name
                string baseName = Path.GetFileNameWithoutExtension("input.xlsx");
                string csvFileName = $"{baseName}_{srcSheet.Name}_Part{partNumber}.csv";

                // Save the chunk as a CSV file
                chunkWorkbook.Save(csvFileName, SaveFormat.Csv);
            }
        }
    }
}
