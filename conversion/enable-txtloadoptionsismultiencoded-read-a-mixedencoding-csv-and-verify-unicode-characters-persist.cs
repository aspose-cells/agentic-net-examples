// Title: Import a CSV with mixed UTF-8 and UTF-16LE lines using Aspose.Cells TxtLoadOptions.IsMultiEncoded and confirm Unicode data integrity in C#
// AI Prompts: Generate C# code that reads a CSV with rows encoded in UTF‑8 and UTF‑16LE by enabling TxtLoadOptions.IsMultiEncoded in Aspose.Cells. | Write a program that loads the mixed‑encoding CSV into a Workbook, extracts the cell contents, and checks that the Chinese and Russian characters are unchanged. | Demonstrate saving the resulting workbook to an XLSX file and printing a confirmation that Unicode data was preserved.
// Common Searches: Aspose.Cells read CSV file that contains both UTF‑8 and UTF‑16LE encoded lines in .NET | C# example using TxtLoadOptions.IsMultiEncoded to import a multi‑encoded CSV | how to ensure Unicode characters stay intact after loading a CSV with Aspose.Cells | convert a CSV with different line encodings to XLSX using Aspose.Cells | validate cell values after importing a CSV with mixed encodings in C#
// Tags: Aspose.Cells TxtLoadOptions multi-encoding CSV import | C# mixed UTF-8 UTF-16LE CSV loading | Unicode data integrity during CSV to workbook conversion | export workbook to XLSX after multi-encoding CSV read | cell value verification after CSV import

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMixedEncodingDemo
{
    // The example creates a temporary CSV where the first line is UTF‑8 and the second line is UTF‑16LE, enables TxtLoadOptions.IsMultiEncoded to load the file into an Aspose.Cells Workbook, reads cell values to confirm that Chinese and Russian characters are retained, and optionally saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Path for the temporary mixed‑encoding CSV file
            string csvPath = Path.Combine(Path.GetTempPath(), "mixed_encoding.csv");

            // Prepare CSV content:
            // Line 1 – UTF‑8 encoded, contains English and Chinese characters
            // Line 2 – UTF‑16LE (Unicode) encoded, contains Russian and Chinese characters
            string line1 = "Hello,世界";
            string line2 = "Привет,世界";

            // Write the two lines with different encodings to the same file
            using (FileStream fs = new FileStream(csvPath, FileMode.Create, FileAccess.Write))
            {
                // Write first line in UTF‑8
                byte[] utf8Bytes = Encoding.UTF8.GetBytes(line1 + Environment.NewLine);
                fs.Write(utf8Bytes, 0, utf8Bytes.Length);

                // Write second line in UTF‑16LE (Unicode)
                byte[] utf16Bytes = Encoding.Unicode.GetBytes(line2 + Environment.NewLine);
                fs.Write(utf16Bytes, 0, utf16Bytes.Length);
            }

            // Configure TxtLoadOptions to handle multiple encodings
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.IsMultiEncoded = true;          // Enable multi‑encoding support
            loadOptions.Separator = ',';                // CSV separator
            loadOptions.Encoding = Encoding.UTF8;       // Default encoding (used for the first part)

            // Load the CSV file with the specified options
            Workbook workbook = new Workbook(csvPath, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Retrieve and display the loaded values
            string a1 = cells["A1"].StringValue; // Expected: "Hello"
            string b1 = cells["B1"].StringValue; // Expected: "世界"
            string a2 = cells["A2"].StringValue; // Expected: "Привет"
            string b2 = cells["B2"].StringValue; // Expected: "世界"

            Console.WriteLine($"A1: {a1}");
            Console.WriteLine($"B1: {b1}");
            Console.WriteLine($"A2: {a2}");
            Console.WriteLine($"B2: {b2}");

            // Simple verification that Unicode characters are preserved
            bool unicodePreserved = b1 == "世界" && b2 == "世界" && a2 == "Привет";
            Console.WriteLine("Unicode characters preserved: " + unicodePreserved);

            // Optional: Save the workbook to an XLSX file to demonstrate successful round‑trip
            string outputXlsx = Path.Combine(Path.GetTempPath(), "MixedEncodingResult.xlsx");
            workbook.Save(outputXlsx, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {outputXlsx}");

            // Clean up temporary CSV file (optional)
            // File.Delete(csvPath);
        }
    }
}
