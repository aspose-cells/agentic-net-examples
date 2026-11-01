// Title: Extract metadata from a password‑protected Excel workbook and save it as JSON using Aspose.Cells for .NET
// Common Searches: how to read built‑in and custom properties from an encrypted .xlsx with Aspose.Cells C# | Aspose.Cells load password protected workbook and export document properties to JSON | C# code to extract Excel metadata from a workbook secured with a password | using LoadOptions to open encrypted Excel file and get its properties in .NET | serialize Aspose.Cells workbook document properties to a JSON file
// Tags: load encrypted workbook Aspose.Cells C# | extract built‑in document properties Aspose.Cells | export custom Excel properties to JSON | Aspose.Cells LoadOptions password protection | serialize workbook metadata JSON .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example verifies the encrypted Excel file, opens it with LoadOptions and the supplied password, gathers both built‑in and custom document properties into a dictionary, serializes the data to an indented JSON string, and writes the output to a file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook and its password
        string workbookPath = "encrypted.xlsx";
        string password = "yourPassword";

        // Verify that the workbook file exists before attempting to load it
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
            return;
        }

        try
        {
            // Load the encrypted workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // Prepare a dictionary to hold metadata
            var metadata = new Dictionary<string, object>();

            // Extract built‑in document properties
            var builtInProps = workbook.BuiltInDocumentProperties;
            metadata["BuiltInProperties"] = new Dictionary<string, object>
            {
                { "Author", builtInProps.Author },
                { "Title", builtInProps.Title },
                { "Subject", builtInProps.Subject },
                { "Keywords", builtInProps.Keywords },
                { "Comments", builtInProps.Comments },
                { "CreatedTime", builtInProps.CreatedTime },
                // Use LastSavedTime instead of the non‑existent LastModifiedTime
                { "LastSavedTime", builtInProps.LastSavedTime },
                { "Category", builtInProps.Category },
                { "Manager", builtInProps.Manager },
                { "Company", builtInProps.Company }
            };

            // Extract custom document properties, if any
            var customProps = workbook.CustomDocumentProperties;
            var customDict = new Dictionary<string, object>();
            foreach (var prop in customProps) // Use var to avoid type resolution issues
            {
                customDict[prop.Name] = prop.Value;
            }
            metadata["CustomProperties"] = customDict;

            // Serialize metadata to JSON with indentation for readability
            string json = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });

            // Save the JSON to a file
            string jsonPath = "metadata.json";
            File.WriteAllText(jsonPath, json);

            Console.WriteLine($"Metadata extracted and saved to '{jsonPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
