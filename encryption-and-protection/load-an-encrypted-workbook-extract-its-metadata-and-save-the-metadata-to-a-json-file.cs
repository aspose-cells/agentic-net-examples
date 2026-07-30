// Title: Extract Excel Workbook Properties to JSON from an Encrypted File (Aspose.Cells C#)
// Description: C# sample that checks for a password‑protected .xlsx, opens it with Aspose.Cells LoadOptions, reads both built‑in and custom document properties, serializes the collected name/value pairs into indented JSON, and writes the output to a metadata.json file.
// Keywords: Aspose.Cells | C# | encrypted workbook | password protected Excel | document properties extraction | built‑in properties | custom properties | metadata to JSON | LoadOptions | Excel automation
// Common Searches: open password protected Excel with Aspose.Cells C# | export Excel built‑in and custom properties to JSON | read metadata from encrypted .xlsx using Aspose | C# code to extract Excel document properties | save workbook metadata as JSON file
// Developer Intent: Load a protected workbook, retrieve its built‑in and custom properties, and persist the data as a JSON document.
// Use Cases: Compliance audits that require metadata from secured spreadsheets. | Migration of property information from encrypted Excel files to a central JSON store for analytics. | Automated validation of workbook protection settings in CI/CD pipelines.
// AI Prompts: Generate C# code that opens an encrypted .xlsx with a given password using Aspose.Cells and lists all built‑in document properties. | Create a method that returns an indented JSON string containing both built‑in and custom properties of a Workbook object. | Suggest robust error‑handling patterns for loading password‑protected workbooks and writing their metadata to a file.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// C# sample that checks for a password‑protected .xlsx, opens it with Aspose.Cells LoadOptions, reads both built‑in and custom document properties, serializes the collected name/value pairs into indented JSON, and writes the output to a metadata.json file.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string encryptedFilePath = "encrypted.xlsx";

        // Password used to protect the workbook
        string password = "your_password";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(encryptedFilePath))
        {
            Console.WriteLine($"File not found: {encryptedFilePath}");
            return;
        }

        try
        {
            // Load the workbook with the password
            var loadOptions = new LoadOptions { Password = password };
            var workbook = new Workbook(encryptedFilePath, loadOptions);

            // Prepare dictionaries to hold built‑in and custom properties
            var builtInProps = new Dictionary<string, object>();
            var customProps = new Dictionary<string, object>();

            // Extract built‑in document properties
            foreach (var prop in workbook.BuiltInDocumentProperties)
            {
                builtInProps[prop.Name] = prop.Value;
            }

            // Extract custom document properties
            foreach (var prop in workbook.CustomDocumentProperties)
            {
                customProps[prop.Name] = prop.Value;
            }

            // Combine into a single object for JSON serialization
            var allMetadata = new
            {
                BuiltInProperties = builtInProps,
                CustomProperties = customProps
            };

            // Serialize to JSON
            string json = JsonSerializer.Serialize(allMetadata, new JsonSerializerOptions { WriteIndented = true });

            // Save JSON to a file
            string jsonOutputPath = "metadata.json";
            File.WriteAllText(jsonOutputPath, json);

            Console.WriteLine($"Metadata extracted and saved to '{jsonOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
