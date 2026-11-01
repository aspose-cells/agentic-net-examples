// Title: Add and retrieve a custom XML part as workbook metadata using Aspose.Cells for .NET (C#)
// Common Searches: how to embed custom XML metadata into an Excel workbook with Aspose.Cells C# | retrieve custom XML part from an existing .xlsx file using Aspose.Cells .NET API | store XML metadata in a workbook and read it back with Aspose.Cells for .NET | Aspose.Cells CustomXmlParts add example in C# | read custom XML part data from a saved workbook using Aspose.Cells
// Tags: add custom xml part to workbook Aspose.Cells | retrieve custom xml part from workbook Aspose.Cells | custom xml metadata storage in .xlsx using C# | Aspose.Cells CustomXmlParts add method example | Aspose.Cells CustomXmlParts read data example

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCustomXmlExample
{
    // The sample creates a new Workbook, converts an XML metadata string to UTF‑8 bytes, adds it as a custom XML part via Workbook.CustomXmlParts.Add, saves the file, reloads the workbook, accesses the first CustomXmlPart, converts its Data back to a string, and prints the retrieved XML metadata.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // XML content to store as metadata
                const string xmlContent = @"<Metadata>
    <Author>John Doe</Author>
    <Created>2023-10-01</Created>
</Metadata>";

                // Convert XML to byte array
                byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlContent);

                // Add the custom XML part to the workbook (schemaData is optional, pass null)
                workbook.CustomXmlParts.Add(xmlBytes, null);

                // Save the workbook
                const string filePath = "CustomXmlPart.xlsx";
                workbook.Save(filePath);
                Console.WriteLine($"Workbook saved to '{filePath}'.");

                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File '{filePath}' was not found.");
                    return;
                }

                // Load the workbook back
                var loadedWorkbook = new Workbook(filePath);

                // Retrieve the first custom XML part (only one was added)
                if (loadedWorkbook.CustomXmlParts.Count > 0)
                {
                    var retrievedPart = loadedWorkbook.CustomXmlParts[0];
                    string retrievedXml = Encoding.UTF8.GetString(retrievedPart.Data);
                    Console.WriteLine("Retrieved XML Metadata:");
                    Console.WriteLine(retrievedXml);
                }
                else
                {
                    Console.WriteLine("No custom XML parts were found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
