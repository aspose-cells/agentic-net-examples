using System;
using System.IO;
using Aspose.Cells.Metadata;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string workbookPath = "input.xlsx";

        // Create metadata options to load document properties
        MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);

        // Load the workbook metadata (uses the provided WorkbookMetadata(string, MetadataOptions) constructor)
        WorkbookMetadata metadata = new WorkbookMetadata(workbookPath, options);

        // Prepare a string builder for the output
        StringWriter writer = new StringWriter();

        // Retrieve and write built‑in document properties
        writer.WriteLine("Built‑in Document Properties:");
        foreach (var prop in metadata.BuiltInDocumentProperties)
        {
            writer.WriteLine($"{prop.Name}: {prop.Value}");
        }

        writer.WriteLine();

        // Retrieve and write custom document properties
        writer.WriteLine("Custom Document Properties:");
        foreach (var prop in metadata.CustomDocumentProperties)
        {
            writer.WriteLine($"{prop.Name}: {prop.Value}");
        }

        // Get the final text
        string result = writer.ToString();

        // Output to console
        Console.WriteLine(result);

        // Save the metadata information to a TXT file
        File.WriteAllText("metadata.txt", result);
    }
}