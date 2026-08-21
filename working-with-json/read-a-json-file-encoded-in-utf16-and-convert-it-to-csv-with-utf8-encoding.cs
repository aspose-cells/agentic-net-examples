// Title: C# – Convert a UTF‑16 JSON file to a UTF‑8 CSV using Aspose.Cells
// Description: Load a UTF‑16 encoded JSON file into an Aspose.Cells Workbook with JsonLoadOptions, then save it as a UTF‑8 encoded CSV using TxtSaveOptions. The example shows stream‑based loading, automatic BOM detection, and encoding configuration for seamless conversion.
// Keywords: Aspose.Cells JSON to CSV | C# UTF-16 JSON read | UTF-8 CSV export | JsonLoadOptions example | TxtSaveOptions encoding | .NET file encoding conversion | Aspose.Cells workbook save CSV
// Common Searches: how to read UTF-16 JSON in C# with Aspose.Cells | convert JSON to CSV with UTF-8 encoding Aspose.Cells | Aspose.Cells load JSON file stream | C# save workbook as CSV UTF-8 | Aspose.Cells encoding options for CSV
// Developer Intent: Read a UTF‑16 JSON file into an Aspose.Cells workbook and export it as a UTF‑8 CSV file.
// Use Cases: Batch conversion of legacy UTF‑16 JSON reports to UTF‑8 CSV for analytics pipelines. | Automated data migration where source systems output UTF‑16 JSON and target systems require UTF‑8 CSV. | Web API that accepts JSON payloads and returns CSV responses with proper Unicode encoding.
// AI Prompts: Write C# code to convert a large UTF‑16 JSON file to CSV using Aspose.Cells without loading the entire file into memory. | Show how to prepend a custom header row to the CSV output before saving with Aspose.Cells. | Provide error‑handling logic for cases where the JSON structure cannot be mapped to a tabular format.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsJsonToCsv
{
    // Load a UTF‑16 encoded JSON file into an Aspose.Cells Workbook with JsonLoadOptions, then save it as a UTF‑8 encoded CSV using TxtSaveOptions. The example shows stream‑based loading, automatic BOM detection, and encoding configuration for seamless conversion.
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file (UTF‑16 encoded)
            string jsonFilePath = "input.json";

            // Path for the resulting CSV file (UTF‑8 encoded)
            string csvFilePath = "output.csv";

            // Open the JSON file as a stream; the UTF‑16 BOM will be detected automatically
            using (FileStream jsonStream = new FileStream(jsonFilePath, FileMode.Open, FileAccess.Read))
            {
                // Load options for JSON files
                JsonLoadOptions loadOptions = new JsonLoadOptions();

                // Load the JSON content into a workbook
                Workbook workbook = new Workbook(jsonStream, loadOptions);

                // Configure CSV save options with UTF‑8 encoding
                TxtSaveOptions saveOptions = new TxtSaveOptions
                {
                    Encoding = Encoding.UTF8
                };

                // Save the workbook as CSV using the specified encoding
                workbook.Save(csvFilePath, saveOptions);
            }

            Console.WriteLine($"JSON file '{jsonFilePath}' has been converted to CSV '{csvFilePath}' with UTF‑8 encoding.");
        }
    }
}
