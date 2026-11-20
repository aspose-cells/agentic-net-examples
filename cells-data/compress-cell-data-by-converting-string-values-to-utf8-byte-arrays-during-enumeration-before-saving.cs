// Title: Compress string values in Excel cells to GZip‑encoded Base64 using Aspose.Cells for .NET
// Common Searches: how to gzip compress string cells in an Excel file with Aspose.Cells C# | store compressed UTF‑8 data as Base64 in worksheet cells using Aspose.Cells | enumerate used cells and replace text with compressed Base64 strings in .NET | Aspose.Cells example for compressing cell content before saving workbook | C# code to convert Excel cell string to UTF‑8 bytes, gzip them, and write back as Base64
// Tags: gzip compress cell string values Aspose.Cells | store compressed data as base64 in Excel cells | convert cell text to utf8 byte array C# | iterate used cells and replace values Aspose.Cells | compress worksheet data before saving workbook

using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells;

// The example creates a workbook, fills several cells with text, iterates over all used cells, and for each string cell converts the text to a UTF‑8 byte array, compresses it with GZip, encodes the compressed bytes to a Base64 string, writes the Base64 string back into the cell, and finally saves the workbook as 'CompressedData.xlsx'.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the provided create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (string values)
        sheet.Cells["A1"].PutValue("Hello World");
        sheet.Cells["A2"].PutValue("Aspose.Cells");
        sheet.Cells["B1"].PutValue("Sample Text");
        sheet.Cells["B2"].PutValue("Compression Test");

        // Enumerate used cells
        Cells cells = sheet.Cells;
        int maxRow = cells.MaxDataRow;
        int maxColumn = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxColumn; col++)
            {
                Cell cell = cells[row, col];

                // Process only string cells
                if (cell.Type == CellValueType.IsString)
                {
                    string original = cell.StringValue;

                    // Convert string to UTF‑8 byte array
                    byte[] utf8Bytes = Encoding.UTF8.GetBytes(original);

                    // Compress the UTF‑8 bytes using GZip
                    byte[] compressedBytes;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (GZipStream gzip = new GZipStream(ms, CompressionMode.Compress, true))
                        {
                            gzip.Write(utf8Bytes, 0, utf8Bytes.Length);
                        }
                        compressedBytes = ms.ToArray();
                    }

                    // Store compressed data back as a Base64 string (still a string value)
                    string base64Compressed = Convert.ToBase64String(compressedBytes);
                    cell.PutValue(base64Compressed);
                }
            }
        }

        // Save the workbook (using the provided save rule)
        workbook.Save("CompressedData.xlsx");
    }
}
