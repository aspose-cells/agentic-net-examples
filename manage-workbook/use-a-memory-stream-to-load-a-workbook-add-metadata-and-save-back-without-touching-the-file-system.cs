// Title: Add Custom Document Properties to an Aspose.Cells Workbook Using In‑Memory Streams (C#)
// Description: Demonstrates how to create a workbook, save it to a MemoryStream, load WorkbookMetadata, add custom document properties, and write the updated metadata back to another MemoryStream—all without touching the file system.
// Keywords: Aspose.Cells | WorkbookMetadata | MemoryStream | custom document properties | C# | .NET | in‑memory Excel | no file I/O | metadata manipulation | SaveFormat.Xlsx
// Common Searches: Aspose.Cells add custom properties without saving to disk | C# load Excel workbook from MemoryStream and edit metadata | How to use WorkbookMetadata with MemoryStream | Update Excel custom document properties in memory | Aspose.Cells metadata stream example
// Developer Intent: Modify or add custom document properties to an Excel workbook entirely in memory and retrieve the resulting metadata stream.
// Use Cases: Generate an Excel report on a web server, embed processing details as custom properties, and return the stream to a client API. | Receive an uploaded Excel file as a stream, append audit information via custom properties, and store the updated stream in a database. | Validate metadata changes in automated tests by inspecting the length or content of the metadata MemoryStream.
// AI Prompts: Show C# code to read existing custom document properties from a workbook loaded from a MemoryStream using Aspose.Cells. | Provide an example that updates built‑in properties (Author, Title) in an in‑memory workbook and returns the modified stream. | Explain how to apply both custom and built‑in property changes to a workbook before saving it to a MemoryStream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

// Demonstrates how to create a workbook, save it to a MemoryStream, load WorkbookMetadata, add custom document properties, and write the updated metadata back to another MemoryStream—all without touching the file system.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello, Aspose!");
        workbook.Worksheets[0].Cells["B1"].PutValue(DateTime.Now);

        // Save the workbook to a memory stream (XLSX format)
        using (MemoryStream workbookStream = new MemoryStream())
        {
            workbook.Save(workbookStream, SaveFormat.Xlsx);
            workbookStream.Position = 0; // Reset for reading

            // Load metadata from the workbook stream
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(workbookStream, metaOptions);

            // Add custom document properties
            metadata.CustomDocumentProperties.Add("ProcessedBy", "AsposeDemo");
            metadata.CustomDocumentProperties.Add("ProcessedOn", DateTime.UtcNow);

            // Save the modified metadata back to a new memory stream
            using (MemoryStream metadataStream = new MemoryStream())
            {
                metadata.Save(metadataStream);
                metadataStream.Position = 0; // Reset if further processing is needed

                // Demonstrate that metadata was saved to the stream
                Console.WriteLine($"Metadata saved to memory stream. Length = {metadataStream.Length} bytes.");
            }
        }

        // Clean up
        workbook.Dispose();
    }
}
