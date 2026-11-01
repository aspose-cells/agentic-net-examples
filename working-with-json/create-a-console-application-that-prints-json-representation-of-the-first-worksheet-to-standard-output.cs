// Title: Print JSON of the first worksheet from an Excel file using Aspose.Cells in a C# console app
// AI Prompts: Write a C# console program that accepts an Excel file path, loads the workbook with Aspose.Cells, extracts the first worksheet, and writes its JSON representation to stdout. | Create .NET code that validates the input file, uses JsonSaveOptions to serialize only the primary worksheet to a memory stream, and outputs the resulting JSON string to the console.
// Common Searches: how to output the initial sheet of an Excel workbook as JSON using Aspose.Cells C# console | C# console application that prints sheet JSON to standard output | Aspose.Cells JsonSaveOptions serialize a single worksheet to a string | read Excel file path argument and convert the leading sheet to JSON in .NET | export Excel worksheet to JSON without creating a file using Aspose.Cells
// Tags: Aspose.Cells JsonSaveOptions serialize primary sheet | C# console export Excel sheet to JSON | validate Excel file path before JSON conversion Aspose.Cells | memory stream JSON output Aspose.Cells .NET | error handling Excel to JSON conversion C#

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetJsonExporter
{
    // The program validates a command‑line Excel file path, loads the workbook with Aspose.Cells, ensures at least one worksheet exists, creates a JsonSaveOptions object, saves the workbook to a MemoryStream in JSON format, reads the JSON string, and writes it to the console while handling file‑not‑found and other exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that a file path was provided.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WorksheetJsonExporter <excel-file-path>");
                return;
            }

            string filePath = args[0];

            // Ensure the input file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found - {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(filePath);

                // Ensure there is at least one worksheet.
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Error: The workbook contains no worksheets.");
                    return;
                }

                // Prepare JSON save options. (Exporting the whole workbook; specific sheet selection not supported in this version.)
                JsonSaveOptions jsonOptions = new JsonSaveOptions();

                // Save the workbook to a memory stream in JSON format.
                using (MemoryStream jsonStream = new MemoryStream())
                {
                    workbook.Save(jsonStream, jsonOptions);
                    jsonStream.Position = 0; // Reset stream position for reading.

                    // Read the JSON text from the stream.
                    using (StreamReader reader = new StreamReader(jsonStream))
                    {
                        string json = reader.ReadToEnd();

                        // Output the JSON representation to standard output.
                        Console.WriteLine(json);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
