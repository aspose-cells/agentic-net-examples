using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsJsonToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file (encoded in UTF‑16)
            string jsonFilePath = "input.json";

            // Load the JSON file into a workbook using JsonLoadOptions
            // (default options are sufficient for this scenario)
            Workbook workbook = new Workbook(jsonFilePath, new JsonLoadOptions());

            // Configure save options for CSV with UTF‑8 encoding
            TxtSaveOptions csvSaveOptions = new TxtSaveOptions
            {
                Encoding = Encoding.UTF8   // Set output encoding to UTF‑8
            };

            // Save the workbook as a CSV file using the specified options
            string csvFilePath = "output.csv";
            workbook.Save(csvFilePath, csvSaveOptions);

            Console.WriteLine($"JSON file '{jsonFilePath}' has been converted to CSV '{csvFilePath}' with UTF‑8 encoding.");
        }
    }
}