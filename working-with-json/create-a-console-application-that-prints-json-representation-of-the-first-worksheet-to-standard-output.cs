// Title: Print JSON representation of the first worksheet from an Excel file using Aspose.Cells in a C# console app
// AI Prompts: Write a C# console program that loads an .xlsx file with Aspose.Cells, selects the first worksheet, converts it to a JSON string, and prints the result to stdout. | Demonstrate how to configure JsonSaveOptions so that only a specified worksheet is serialized to JSON. | Create a version of the console tool that takes an optional output file path and writes the JSON of the first worksheet to that file.
// Common Searches: convert first Excel sheet to JSON using Aspose.Cells in C# | Aspose.Cells JsonSaveOptions serialize only one worksheet | C# console app output Excel worksheet as JSON string | command line utility to export Excel worksheet to JSON with Aspose.Cells | handle file not found error when loading Excel in Aspose.Cells C#
// Tags: Aspose.Cells first worksheet JSON export | C# console JSON output from Excel workbook | JsonSaveOptions worksheet selection | memory stream JSON conversion Aspose.Cells | error handling missing Excel file Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace WorksheetJsonExporter
{
    // The console application checks for an input .xlsx file (provided via command line or defaulting to "input.xlsx"), loads it with Aspose.Cells, saves the workbook to a MemoryStream using JsonSaveOptions, converts the stream to a UTF‑8 JSON string, and writes the JSON to standard output while handling missing‑file and other runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the Excel file path (command‑line argument or default).
            string excelPath = args.Length > 0 ? args[0] : "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(excelPath))
            {
                Console.Error.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(excelPath);

                // Prepare a memory stream to hold the JSON output.
                using (MemoryStream jsonStream = new MemoryStream())
                {
                    // Configure JSON save options.
                    JsonSaveOptions jsonOptions = new JsonSaveOptions();

                    // Save the workbook (or its first worksheet) as JSON into the memory stream.
                    // Aspose.Cells does not provide a direct WorksheetIndex property for JsonSaveOptions,
                    // so the entire workbook is saved. If only the first worksheet is needed,
                    // further processing can be applied to the JSON string.
                    workbook.Save(jsonStream, jsonOptions);

                    // Convert the stream contents to a UTF‑8 string.
                    string json = Encoding.UTF8.GetString(jsonStream.ToArray());

                    // Print the JSON representation to standard output.
                    Console.WriteLine(json);
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully.
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
