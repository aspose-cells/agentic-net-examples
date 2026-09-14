// Title: Load an Excel workbook from a byte array, add built‑in and custom document properties, and save back to a byte array using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an Excel workbook from a byte array using a MemoryStream, sets the Title, Author, Company and two custom properties, then outputs the revised workbook as a byte array with Aspose.Cells. | Write a method that accepts workbook bytes, injects built‑in metadata and user‑defined properties, and returns the updated bytes without writing any files, using Aspose.Cells for .NET. | Provide a sample that demonstrates modifying Excel document metadata entirely in memory and retrieving the resulting byte[] with Aspose.Cells.
// Common Searches: Aspose.Cells add custom document properties to workbook from memory stream | How to modify Excel metadata in C# without saving to disk | Load Excel from byte array, set title and author, and get updated bytes using Aspose.Cells | Save modified workbook to byte array using Aspose.Cells for .NET
// Tags: load workbook from memory stream Aspose.Cells | set built‑in document properties Aspose.Cells | add custom document properties Aspose.Cells | save workbook to byte array .NET | modify Excel metadata in memory

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook from a byte array via MemoryStream, updates built‑in properties (Title, Author, Company) and custom properties (ProjectId, Reviewed), then saves the workbook to another MemoryStream and returns the updated byte array for further processing.
class Program
{
    static void Main()
    {
        // Obtain the original workbook bytes from any source (e.g., database, API)
        byte[] originalBytes = GetWorkbookBytes();

        // Load the workbook from a memory stream
        using (MemoryStream inputStream = new MemoryStream(originalBytes))
        {
            Workbook workbook = new Workbook(inputStream);

            // Add built‑in metadata
            workbook.BuiltInDocumentProperties.Title = "Sales Report";
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Company = "Acme Corp";

            // Add custom metadata
            workbook.CustomDocumentProperties.Add("ProjectId", "12345");
            workbook.CustomDocumentProperties.Add("Reviewed", true);

            // Save the modified workbook back to a new memory stream
            using (MemoryStream outputStream = new MemoryStream())
            {
                workbook.Save(outputStream, SaveFormat.Xlsx);

                // The updated workbook bytes are now in outputStream
                byte[] updatedBytes = outputStream.ToArray();

                // Use the updated bytes as needed (e.g., send over network, store in DB)
                ProcessUpdatedWorkbook(updatedBytes);
            }
        }
    }

    // Example method to provide initial workbook bytes
    static byte[] GetWorkbookBytes()
    {
        // Create a simple workbook in memory and return its bytes
        using (MemoryStream ms = new MemoryStream())
        {
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Hello World");
            wb.Save(ms, SaveFormat.Xlsx);
            return ms.ToArray();
        }
    }

    // Example method to handle the updated workbook bytes
    static void ProcessUpdatedWorkbook(byte[] data)
    {
        // For demonstration, output the size of the updated workbook
        Console.WriteLine($"Updated workbook size: {data.Length} bytes");
    }
}
