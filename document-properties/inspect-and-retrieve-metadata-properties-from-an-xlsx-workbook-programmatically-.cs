using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Path to the existing XLSX workbook
        string filePath = "Sample.xlsx";

        // Create options to load only document properties metadata
        MetadataOptions options = new MetadataOptions(MetadataType.DocumentProperties);

        // Load the workbook metadata using the provided constructor
        WorkbookMetadata metadata = new WorkbookMetadata(filePath, options);

        // Access built‑in document properties (read‑only collection)
        BuiltInDocumentPropertyCollection builtInProps = metadata.BuiltInDocumentProperties;
        Console.WriteLine("Built‑in Document Properties:");
        Console.WriteLine($"Author: {builtInProps.Author}");
        Console.WriteLine($"Title: {builtInProps.Title}");
        Console.WriteLine($"Subject: {builtInProps.Subject}");
        Console.WriteLine($"Company: {builtInProps.Company}");
        Console.WriteLine($"Created Time: {builtInProps.CreatedTime}");

        // Access custom document properties (read‑write collection)
        CustomDocumentPropertyCollection customProps = metadata.CustomDocumentProperties;
        Console.WriteLine("\nCustom Document Properties:");
        foreach (DocumentProperty prop in customProps)
        {
            Console.WriteLine($"{prop.Name}: {prop.Value}");
        }
    }
}