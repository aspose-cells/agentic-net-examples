// Title: C# – Load an encrypted Excel workbook with Aspose.Cells, extract built‑in and custom document properties, and save them as JSON
// Common Searches: C# load password protected .xlsx using Aspose.Cells LoadOptions | extract built‑in document properties from an encrypted Excel file with Aspose.Cells | export custom Excel document properties to a JSON file in .NET | retrieve workbook metadata after opening an encrypted workbook in Aspose.Cells | serialize Aspose.Cells workbook properties to JSON with System.Text.Json
// Tags: Aspose.Cells load encrypted workbook | extract built‑in document properties C# | custom document properties to JSON | password protected Excel metadata extraction | Workbook BuiltInDocumentProperties Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example demonstrates how to open a password‑protected Excel workbook using Aspose.Cells LoadOptions, collect both built‑in and custom document properties into a dictionary, serialize the dictionary to a formatted JSON string with System.Text.Json, and write the JSON to a file, handling missing files and load errors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the encrypted workbook and its password
            string workbookPath = "encrypted.xlsx";
            string password = "myPassword";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
                return;
            }

            // Load the encrypted workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // Prepare a dictionary to hold metadata (built‑in + custom)
            var metadata = new Dictionary<string, object>();

            // ----- Built‑in document properties -----
            // Aspose.Cells stores built‑in properties in the BuiltInDocumentProperties collection
            var builtInProps = workbook.BuiltInDocumentProperties;

            // Add selected built‑in properties to the dictionary
            metadata["Title"] = builtInProps.Title;
            metadata["Subject"] = builtInProps.Subject;
            metadata["Author"] = builtInProps.Author;
            metadata["Keywords"] = builtInProps.Keywords;
            metadata["Comments"] = builtInProps.Comments;
            metadata["LastSavedBy"] = builtInProps.LastSavedBy;          // corrected property name
            metadata["CreatedTime"] = builtInProps.CreatedTime;
            metadata["LastSavedTime"] = builtInProps.LastSavedTime;    // corrected property name
            metadata["Category"] = builtInProps.Category;
            metadata["Manager"] = builtInProps.Manager;
            metadata["Company"] = builtInProps.Company;

            // ----- Custom document properties -----
            // Custom properties are stored in the CustomDocumentProperties collection
            var customProps = workbook.CustomDocumentProperties;
            foreach (var prop in customProps)
            {
                // Use the property name as the key and its value as the dictionary value
                metadata[prop.Name] = prop.Value;
            }

            // Serialize the metadata dictionary to a formatted JSON string
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
