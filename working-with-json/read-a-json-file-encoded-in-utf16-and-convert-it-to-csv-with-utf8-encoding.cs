using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file (UTF‑16 encoded)
            string jsonFilePath = "input.json";

            // Load JSON with default options (keeps schema if needed)
            JsonLoadOptions jsonLoadOptions = new JsonLoadOptions
            {
                KeptSchema = true
            };
            Workbook workbook = new Workbook(jsonFilePath, jsonLoadOptions);

            // Configure CSV (TXT) save options with UTF‑8 encoding
            TxtSaveOptions csvSaveOptions = new TxtSaveOptions
            {
                Encoding = Encoding.UTF8,
                Separator = ',' // ensure comma delimiter
            };

            // Save the workbook as CSV (UTF‑8)
            string csvFilePath = "output.csv";
            workbook.Save(csvFilePath, csvSaveOptions);

            Console.WriteLine($"JSON file '{jsonFilePath}' has been converted to CSV '{csvFilePath}' with UTF‑8 encoding.");
        }
    }
}