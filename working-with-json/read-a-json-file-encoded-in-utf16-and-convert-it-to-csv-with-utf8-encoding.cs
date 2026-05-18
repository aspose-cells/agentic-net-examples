using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsJsonToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file (UTF‑16 encoded)
            string jsonFilePath = "input.json";

            // Read the raw bytes of the UTF‑16 file
            byte[] jsonBytes = File.ReadAllBytes(jsonFilePath);

            // Load options for JSON files
            JsonLoadOptions loadOptions = new JsonLoadOptions();

            // Load the JSON content into a workbook using a memory stream
            Workbook workbook = new Workbook(new MemoryStream(jsonBytes), loadOptions);

            // Configure save options for CSV with UTF‑8 encoding
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Encoding = Encoding.UTF8,   // Desired UTF‑8 output
                Separator = ','            // Use comma as CSV delimiter
            };

            // Save the workbook as CSV
            string csvOutputPath = "output.csv";
            workbook.Save(csvOutputPath, saveOptions);

            Console.WriteLine($"JSON file '{jsonFilePath}' has been converted to UTF‑8 CSV at '{csvOutputPath}'.");
        }
    }
}