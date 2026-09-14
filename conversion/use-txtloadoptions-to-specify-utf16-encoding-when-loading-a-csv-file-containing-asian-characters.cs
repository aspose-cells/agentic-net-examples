// Title: Read a UTF-16 CSV containing Japanese and Korean characters using TxtLoadOptions in Aspose.Cells for .NET
// AI Prompts: Configure TxtLoadOptions with Encoding.Unicode, load the CSV file, and output the value of cell A1. | Generate a Unicode-encoded CSV with Japanese and Korean text, then open it using Aspose.Cells specifying the Unicode encoding to verify correct rendering.
// Common Searches: Aspose.Cells C# load CSV file with UTF-16 encoding for Asian characters | How to specify Unicode encoding when importing a CSV into Aspose.Cells workbook | Reading Japanese and Korean text from a UTF-16 CSV using Aspose.Cells TxtLoadOptions
// Tags: Custom text encoding for CSV import | Aspose.Cells CSV loading with specific encoding | C# handling Asian characters in CSV files | Workbook load options for text files

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsUtf16Example
{
    // The example creates a UTF-16 CSV file containing Japanese and Korean text if it does not exist, sets TxtLoadOptions.Encoding to Encoding.Unicode, loads the file into an Aspose.Cells Workbook, and prints the value of cell A1 to confirm that Asian characters are read correctly.
    class Program
    {
        static void Main()
        {
            // Path to the CSV file
            string csvPath = "asian_utf16.csv";

            // Ensure the CSV file exists; create a sample UTF‑16 file if missing
            if (!File.Exists(csvPath))
            {
                string[] sampleLines =
                {
                    "こんにちは,世界",   // Japanese
                    "안녕하세요,세계"      // Korean
                };
                // Write the sample lines using UTF‑16 (Unicode) encoding
                File.WriteAllLines(csvPath, sampleLines, Encoding.Unicode);
            }

            try
            {
                // Set load options to use UTF‑16 encoding for text files
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    Encoding = Encoding.Unicode // UTF‑16 LE
                };

                // Load the CSV file with the specified options
                Workbook workbook = new Workbook(csvPath, loadOptions);

                // Access the first worksheet and its cells
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Display a sample cell value to verify correct loading
                Console.WriteLine("Cell A1 value: " + cells["A1"].StringValue);
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during loading or processing
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
