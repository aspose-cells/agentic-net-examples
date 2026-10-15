// Title: Compress string cell values to GZip‑encoded Base64 in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates every string cell, compresses its UTF‑8 bytes using GZip, encodes the result to Base64, and writes it back to the cell. | Create a .NET console utility that reads a workbook, enumerates the used range of each worksheet, applies GZip compression to textual cell contents, stores the compressed data as a Base64 string, and saves the modified file. | Adapt an existing Aspose.Cells example to replace each string cell value with its GZip‑compressed Base64 representation while preserving the workbook structure.
// Common Searches: how to gzip compress text in Excel cells using Aspose.Cells C# | store compressed string as base64 in Excel worksheet with Aspose.Cells .NET | enumerate used range and compress string values in Aspose.Cells workbook | C# example for compressing cell data before saving Excel file
// Tags: gzip compression of string cells Aspose.Cells | utf-8 byte array conversion for Excel cells | base64 encoding of compressed cell values | used-range enumeration Aspose.Cells .NET | c# utility to compress Excel cell text

using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells;

// Loads input.xlsx, iterates each worksheet's used range, compresses every string cell's text to UTF‑8, GZip‑compresses it, encodes the bytes as Base64, writes the result back to the cell, and saves the workbook as output.xlsx.
class CompressCellData
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Get the used range of the worksheet (Aspose.Cells.Range object)
                Aspose.Cells.Range usedRange = cells.MaxDisplayRange;

                // Determine start and end indices
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Enumerate each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only string values
                        if (cell.Type == CellValueType.IsString)
                        {
                            string originalText = cell.StringValue;

                            // Convert the string to UTF‑8 byte array
                            byte[] utf8Bytes = Encoding.UTF8.GetBytes(originalText);

                            // Compress the UTF‑8 bytes using GZip
                            byte[] compressedBytes;
                            using (MemoryStream output = new MemoryStream())
                            {
                                using (GZipStream gzip = new GZipStream(output, CompressionMode.Compress, leaveOpen: true))
                                {
                                    gzip.Write(utf8Bytes, 0, utf8Bytes.Length);
                                }
                                compressedBytes = output.ToArray();
                            }

                            // Store the compressed data as a Base64 string back into the cell
                            string base64Compressed = Convert.ToBase64String(compressedBytes);
                            cell.PutValue(base64Compressed);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
