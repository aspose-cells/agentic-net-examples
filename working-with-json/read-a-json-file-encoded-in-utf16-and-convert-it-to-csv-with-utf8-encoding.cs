// Title: Read a UTF‑16 JSON file and export it to a UTF‑8 CSV using Aspose.Cells for .NET (C#)
// Common Searches: c# read utf-16 json file with aspose.cells and save as utf-8 csv | how to load json data into an Aspose.Cells workbook from a memory stream | aspose.cells txtsaveoptions example for csv with utf-8 encoding | convert unicode json to csv using Aspose.Cells in .NET | aspose.cells loadoptions json format c# sample
// Tags: Aspose.Cells LoadOptions JSON | Aspose.Cells TxtSaveOptions CSV UTF-8 | C# read UTF-16 JSON into workbook | export workbook to CSV with UTF-8 encoding | memory stream JSON loading Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample reads a UTF‑16 (Unicode) JSON file, loads its content into an Aspose.Cells Workbook via a memory stream using LoadOptions set to JSON, and then saves the workbook as a CSV file with UTF‑8 encoding configured through TxtSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for input JSON and output CSV
            string inputPath = "input.json";
            string outputPath = "output.csv";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Read the JSON file encoded in UTF‑16 (Unicode)
            string jsonContent = File.ReadAllText(inputPath, Encoding.Unicode);

            // Load JSON data into a workbook using LoadOptions
            Workbook workbook;
            byte[] jsonBytes = Encoding.Unicode.GetBytes(jsonContent);
            using (MemoryStream ms = new MemoryStream(jsonBytes))
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);
                workbook = new Workbook(ms, loadOptions);
            }

            // Configure CSV save options to use UTF‑8 encoding
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8
            };

            // Save the workbook as a CSV file with UTF‑8 encoding
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"CSV file saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
