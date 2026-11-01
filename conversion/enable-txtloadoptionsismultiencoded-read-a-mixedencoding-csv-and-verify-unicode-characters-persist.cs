// Title: Read a mixed‑encoding CSV (UTF‑8 + UTF‑16) with Aspose.Cells TxtLoadOptions.IsMultiEncoded and preserve Unicode characters in C#
// Common Searches: Aspose.Cells read CSV file that contains both UTF-8 and UTF-16 lines in C# | Enable TxtLoadOptions.IsMultiEncoded to handle multiple encodings in a CSV with Aspose.Cells | Verify Unicode characters are preserved when loading a mixed‑encoding CSV using Aspose.Cells .NET | How to load a CSV with mixed BOMs (UTF-8 and UTF-16LE) using Aspose.Cells TxtLoadOptions | C# example for detecting multiple encodings in a CSV with Aspose.Cells
// Tags: TxtLoadOptions.IsMultiEncoded CSV import | Aspose.Cells multi‑encoding CSV detection | preserve Unicode cells when loading CSV | C# load UTF-8 and UTF-16 CSV with Aspose | mixed‑BOM CSV handling Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample creates a CSV where the first line is UTF‑8 (with BOM) and the second line is UTF‑16LE (with BOM), enables TxtLoadOptions.IsMultiEncoded to automatically detect both encodings, loads the file into an Aspose.Cells Workbook, reads the Unicode values from cells A2 and B2 to confirm they are intact, and optionally saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        // Path for the mixed‑encoding CSV file
        string csvPath = "mixed.csv";

        // Create a CSV file that contains UTF‑8 encoded line followed by a UTF‑16LE encoded line
        using (FileStream fs = new FileStream(csvPath, FileMode.Create, FileAccess.Write))
        {
            // Write UTF‑8 BOM and a header line
            fs.Write(new byte[] { 0xEF, 0xBB, 0xBF }, 0, 3);
            string header = "Name,Comment\n";
            byte[] headerBytes = Encoding.UTF8.GetBytes(header);
            fs.Write(headerBytes, 0, headerBytes.Length);

            // Write UTF‑16LE BOM and a data line with Unicode characters
            fs.Write(new byte[] { 0xFF, 0xFE }, 0, 2);
            string dataLine = "张三,こんにちは\n";
            byte[] dataBytes = Encoding.Unicode.GetBytes(dataLine);
            fs.Write(dataBytes, 0, dataBytes.Length);
        }

        // Enable multi‑encoding detection while loading the CSV
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            IsMultiEncoded = true // Detect and handle multiple encodings in the same file
        };

        // Load the CSV into a workbook using the specified options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve cell values to verify Unicode characters are preserved
        string name = sheet.Cells["A2"].StringValue;      // Expected: 张三
        string comment = sheet.Cells["B2"].StringValue;  // Expected: こんにちは

        // Output the results
        Console.WriteLine($"A2 (Name): {name}");
        Console.WriteLine($"B2 (Comment): {comment}");

        // Save as an Excel file to confirm successful import (optional)
        workbook.Save("output.xlsx");
    }
}
