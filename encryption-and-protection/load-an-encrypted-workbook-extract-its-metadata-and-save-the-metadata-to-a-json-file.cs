// Title: Read encrypted Excel workbook metadata and export to JSON with Aspose.Cells for .NET
// Description: C# sample that opens a password‑protected .xlsx using Aspose.Cells MetadataOptions, extracts both built‑in and custom document properties via WorkbookMetadata, and writes the collected data to a formatted JSON file.
// Keywords: Aspose.Cells encrypted workbook | C# read password protected Excel metadata | WorkbookMetadata password option | extract document properties Aspose.Cells | save Excel metadata to JSON | MetadataOptions example .NET | Excel file encryption Aspose | built‑in and custom properties extraction | Aspose.Cells JSON serialization
// Common Searches: How to read metadata from a password‑protected Excel file using Aspose.Cells | Aspose.Cells C# extract document properties from encrypted workbook | Export Excel workbook properties to JSON in .NET | Load encrypted workbook metadata without opening the file | Aspose.Cells MetadataOptions password example
// Developer Intent: Open a password‑protected Excel workbook, retrieve its built‑in and custom document properties, and save the information as a JSON file.
// Use Cases: Create compliance reports by exporting properties of protected workbooks to a central JSON store. | Maintain an audit trail of encrypted Excel files by archiving their metadata. | Migrate custom and built‑in properties from password‑locked spreadsheets to a metadata management system.
// AI Prompts: Generate C# code that uses Aspose.Cells to open a password‑protected .xlsx and collect all document properties into dictionaries. | Show how to serialize the extracted workbook metadata to a pretty‑printed JSON file with System.Text.Json. | Explain best practices for handling FileNotFoundException, CellsException, and generic errors when loading encrypted workbooks with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

// C# sample that opens a password‑protected .xlsx using Aspose.Cells MetadataOptions, extracts both built‑in and custom document properties via WorkbookMetadata, and writes the collected data to a formatted JSON file.
class ExtractEncryptedWorkbookMetadata
{
    static void Main()
    {
        // Path to the encrypted workbook and its password
        string workbookPath = "encrypted_workbook.xlsx";
        string workbookPassword = "mySecretPassword";

        try
        {
            // Ensure the workbook exists; if not, create a simple encrypted workbook for demo purposes
            if (!File.Exists(workbookPath))
            {
                // Create a new workbook with a single sheet and some data
                var wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Sample Data");

                // Set workbook password for encryption
                wb.Settings.Password = workbookPassword;

                // Save the encrypted workbook
                wb.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Created encrypted workbook at '{workbookPath}'.");
            }

            // Configure metadata options to load document properties and provide the password
            var metadataOptions = new MetadataOptions(MetadataType.DocumentProperties)
            {
                Password = workbookPassword
            };

            // Load the workbook metadata using the provided constructor
            var metadata = new WorkbookMetadata(workbookPath, metadataOptions);

            // Prepare containers for built‑in and custom properties
            var builtInProps = new Dictionary<string, object>();
            var customProps = new Dictionary<string, object>();

            // Extract built‑in document properties
            foreach (DocumentProperty prop in metadata.BuiltInDocumentProperties)
            {
                builtInProps[prop.Name] = prop.Value ?? string.Empty;
            }

            // Extract custom document properties
            foreach (DocumentProperty prop in metadata.CustomDocumentProperties)
            {
                customProps[prop.Name] = prop.Value ?? string.Empty;
            }

            // Combine both sections into a single object for JSON serialization
            var allMetadata = new
            {
                BuiltIn = builtInProps,
                Custom = customProps
            };

            // Serialize to JSON (indented for readability)
            string json = JsonSerializer.Serialize(allMetadata, new JsonSerializerOptions { WriteIndented = true });

            // Save the JSON to a file
            string jsonPath = "workbook_metadata.json";
            File.WriteAllText(jsonPath, json);

            Console.WriteLine($"Metadata extracted and saved to '{jsonPath}'.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File not found: {fnfEx.FileName}");
        }
        catch (CellsException cellsEx)
        {
            Console.WriteLine($"Aspose.Cells error: {cellsEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
