// Title: Convert a UTF-16 encoded JSON file to a UTF-8 CSV file using Aspose.Cells in C#
// AI Prompts: Read a UTF-16 JSON file, load its content into an Aspose.Cells Workbook via a MemoryStream, and export the workbook as a CSV with UTF-8 encoding. | Configure LoadOptions with LoadFormat.Json to import JSON data and set TxtSaveOptions.Encoding to Encoding.UTF8 when saving the workbook as CSV. | Validate the source JSON path, create the destination folder if it does not exist, and handle any exceptions that occur during conversion.
// Common Searches: how to import a UTF-16 JSON file into Aspose.Cells workbook c# | aspocells save workbook as CSV with UTF-8 encoding | c# convert json data to csv using Aspose.Cells memory stream | handle missing input file when converting JSON to CSV in .NET
// Tags: json to workbook conversion Aspose.Cells | csv export with UTF-8 encoding Aspose.Cells | utf-16 file reading c# | memory stream usage for JSON Aspose.Cells | ensure output folder exists c#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The program reads a UTF-16 encoded JSON file, loads it into an Aspose.Cells Workbook via a MemoryStream, and saves the workbook as a UTF-8 encoded CSV file, creating the output directory if necessary and handling missing‑file errors.
class JsonToCsvConverter
{
    static void Main()
    {
        try
        {
            // Path to the UTF‑16 encoded JSON file
            string jsonFilePath = "input.json";

            // Verify that the JSON file exists
            if (!File.Exists(jsonFilePath))
            {
                Console.WriteLine($"Error: JSON file not found at path '{jsonFilePath}'.");
                return;
            }

            // Read the JSON content using UTF‑16 (Unicode) encoding
            string jsonContent = File.ReadAllText(jsonFilePath, Encoding.Unicode);

            // Load the JSON data into a workbook using a memory stream and JSON load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);
            using (MemoryStream ms = new MemoryStream(Encoding.Unicode.GetBytes(jsonContent)))
            {
                Workbook workbook = new Workbook(ms, loadOptions);

                // Define CSV save options with UTF‑8 encoding
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Encoding = Encoding.UTF8
                };

                // Path for the output CSV file
                string csvFilePath = "output.csv";

                // Ensure the output directory exists
                string csvDir = Path.GetDirectoryName(csvFilePath);
                if (!string.IsNullOrEmpty(csvDir) && !Directory.Exists(csvDir))
                {
                    Directory.CreateDirectory(csvDir);
                }

                // Save the workbook as a CSV file with UTF‑8 encoding
                workbook.Save(csvFilePath, csvOptions);
            }

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
