// Title: Read a password‑protected Excel workbook, extract built‑in and custom document properties, and export them to a formatted JSON file using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells LoadOptions with a password to open a protected Excel workbook, retrieve all built‑in and custom document properties from the Workbook object, and serialize the collected data to an indented JSON string. | Write C# code that checks for the source file, creates the output directory if needed, and saves the extracted Excel metadata as a pretty‑printed JSON file at a given path.
// Common Searches: how to extract document properties from a password protected Excel file using Aspose.Cells in C# | Aspose.Cells load encrypted workbook and read custom properties | save Excel workbook metadata to JSON with .NET | C# code to open encrypted .xlsx with password and export metadata | retrieve built‑in document properties from a protected workbook using Aspose.Cells
// Tags: Aspose.Cells load encrypted workbook password | extract built-in document properties Aspose.Cells | read custom document properties Excel .NET | serialize workbook metadata to JSON C# | ensure output directory exists before writing JSON

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace WorkbookMetadataExtractor
{
    // The example verifies the encrypted Excel file exists, loads it with the supplied password via LoadOptions, gathers both built‑in and custom document properties from the Workbook, serializes the metadata into an indented JSON string, creates the output folder if necessary, and writes the JSON to the specified file path.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted Excel file
            string inputPath = @"C:\Data\EncryptedWorkbook.xlsx";

            // Password for the encrypted workbook
            string password = "YourPassword";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook with the password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Prepare a dictionary to hold metadata
                var metadata = new Dictionary<string, object>();

                // Extract built‑in document properties via Workbook.BuiltInDocumentProperties
                var props = workbook.BuiltInDocumentProperties;
                metadata["Author"] = props.Author;
                metadata["Title"] = props.Title;
                metadata["Subject"] = props.Subject;
                metadata["Keywords"] = props.Keywords;
                metadata["Comments"] = props.Comments;
                metadata["CreatedTime"] = props.CreatedTime;
                metadata["LastModifiedTime"] = props.LastSavedTime;
                metadata["Company"] = props.Company;

                // Extract custom document properties
                var customDict = new Dictionary<string, object>();
                foreach (var prop in workbook.CustomDocumentProperties)
                {
                    customDict[prop.Name] = prop.Value;
                }
                metadata["CustomProperties"] = customDict;

                // Serialize metadata to JSON
                string jsonOutput = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });

                // Path to the output JSON file
                string outputPath = @"C:\Data\WorkbookMetadata.json";

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save JSON to file
                File.WriteAllText(outputPath, jsonOutput);

                Console.WriteLine("Metadata extracted and saved to JSON successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
