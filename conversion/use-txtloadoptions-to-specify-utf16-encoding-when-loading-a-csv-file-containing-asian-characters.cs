using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadCsvUtf16Demo
    {
        public static void Run()
        {
            try
            {
                // Path to the CSV file saved with UTF‑16 (Unicode) encoding.
                string csvPath = "asian_utf16.csv";

                // Verify that the file exists to avoid FileNotFoundException.
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"File not found: {csvPath}");
                    return;
                }

                // Set load options to use UTF‑16 encoding.
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    Encoding = Encoding.Unicode
                };

                // Load the CSV file with the specified options.
                Workbook workbook = new Workbook(csvPath, loadOptions);
                Worksheet sheet = workbook.Worksheets[0];

                // Display a cell value to verify correct loading.
                Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadCsvUtf16Demo.Run();
        }
    }
}