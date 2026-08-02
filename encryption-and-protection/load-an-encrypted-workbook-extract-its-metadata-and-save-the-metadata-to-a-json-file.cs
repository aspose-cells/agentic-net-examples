// Title: Extract metadata from a password‑protected Excel workbook to JSON with Aspose.Cells for .NET
// Description: This example demonstrates how to open an encrypted .xlsx file using Aspose.Cells' MetadataOptions, read both built‑in and custom document properties via WorkbookMetadata, serialize the combined data with System.Text.Json, and write the formatted JSON to disk. Includes robust error handling for CellsException and general exceptions.
// Keywords: Aspose.Cells encrypted workbook metadata | C# read password protected Excel properties | WorkbookMetadata document properties | export Excel metadata to JSON | .NET extract custom document properties | MetadataOptions password
// Common Searches: load password protected Excel file Aspose.Cells C# | extract built‑in and custom properties from encrypted workbook | save Excel workbook metadata as JSON .NET | Aspose.Cells MetadataOptions example | read document properties from protected .xlsx
// Developer Intent: Open a password‑protected Excel file, retrieve its built‑in and custom document properties, and save the information as a JSON document.
// Use Cases: Compliance audit: pull metadata from secured workbooks without manual decryption. | Data migration: move Excel document properties into a JSON‑based configuration repository. | Reporting: generate a consolidated JSON report of properties across multiple protected files.
// AI Prompts: Generate C# code that opens an encrypted Excel workbook with Aspose.Cells, reads all document properties, and writes them to a pretty‑printed JSON file. | Explain the role of MetadataOptions.Password when accessing protected workbook metadata in Aspose.Cells for .NET. | Suggest best‑practice error handling for extracting metadata from a password‑protected workbook using Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

namespace AsposeCellsMetadataExample
{
    // This example demonstrates how to open an encrypted .xlsx file using Aspose.Cells' MetadataOptions, read both built‑in and custom document properties via WorkbookMetadata, serialize the combined data with System.Text.Json, and write the formatted JSON to disk. Includes robust error handling for CellsException and general exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the encrypted workbook
                string workbookPath = "encrypted_workbook.xlsx";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file '{workbookPath}' not found.");
                    return;
                }

                // Password used to protect the workbook
                string password = "mySecretPassword";

                // Create options to load document properties and provide the password
                MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties)
                {
                    Password = password
                };

                // Load workbook metadata (WorkbookMetadata does not require explicit Load call)
                WorkbookMetadata metadata = new WorkbookMetadata(workbookPath, options);

                // Extract built‑in document properties
                BuiltInDocumentPropertyCollection builtInProps = metadata.BuiltInDocumentProperties;
                var builtInDict = new Dictionary<string, object>();
                foreach (DocumentProperty prop in builtInProps)
                {
                    builtInDict[prop.Name] = prop.Value;
                }

                // Extract custom document properties
                CustomDocumentPropertyCollection customProps = metadata.CustomDocumentProperties;
                var customDict = new Dictionary<string, object>();
                foreach (DocumentProperty prop in customProps)
                {
                    customDict[prop.Name] = prop.Value;
                }

                // Combine both collections into a single object for JSON serialization
                var metadataObject = new
                {
                    BuiltInProperties = builtInDict,
                    CustomProperties = customDict
                };

                // Serialize to JSON
                string json = JsonSerializer.Serialize(metadataObject, new JsonSerializerOptions { WriteIndented = true });

                // Save JSON to a file
                string jsonPath = "workbook_metadata.json";
                File.WriteAllText(jsonPath, json);

                Console.WriteLine($"Metadata extracted and saved to '{jsonPath}'.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
