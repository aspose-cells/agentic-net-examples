// Title: Add custom and built‑in document properties to an Excel workbook using Aspose.Cells with MemoryStream in C#
// Common Searches: how to add custom document properties to an Excel workbook with Aspose.Cells without writing to a file | load an Excel file from a MemoryStream, modify its metadata, and get a byte array using Aspose.Cells C# | set the built‑in Title property of a workbook loaded from a stream with Aspose.Cells | save an updated Excel workbook to a MemoryStream instead of the file system using Aspose.Cells .NET | Aspose.Cells example for reading and writing Excel workbooks entirely in memory in C#
// Tags: Aspose.Cells memory stream workbook handling | Aspose.Cells custom metadata API | Aspose.Cells change workbook title | Aspose.Cells save workbook to byte array | Aspose.Cells load workbook from stream

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create an Excel workbook, save it to a MemoryStream, reload it, add a custom document property and a built‑in Title property, and then save the modified workbook back to another MemoryStream, returning the bytes—all without accessing the file system, using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create an initial workbook and store it in a memory stream.
        // In a real scenario, the stream could come from a database, network, etc.
        using (var initialStream = new MemoryStream())
        {
            // Create a new workbook with a default worksheet.
            var workbook = new Workbook();
            workbook.Worksheets[0].Name = "Sheet1";

            // Save the workbook into the memory stream.
            workbook.Save(initialStream, SaveFormat.Xlsx);

            // Reset the stream position to the beginning for reading.
            initialStream.Position = 0;

            // Load the workbook from the memory stream.
            var loadedWorkbook = new Workbook(initialStream);

            // ---- Add metadata ----

            // Add a custom document property.
            var customProps = loadedWorkbook.CustomDocumentProperties;
            customProps.Add("Author", "John Doe"); // Example custom property.

            // Set a built‑in document property.
            loadedWorkbook.BuiltInDocumentProperties.Title = "Sample Workbook";

            // ---- Save back to a memory stream ----
            using (var outputStream = new MemoryStream())
            {
                loadedWorkbook.Save(outputStream, SaveFormat.Xlsx);

                // The outputStream now contains the updated workbook bytes.
                // No file system access has occurred.
                byte[] resultBytes = outputStream.ToArray();

                // Example usage of the resulting byte array.
                Console.WriteLine($"Workbook saved to memory. Size = {resultBytes.Length} bytes");
            }
        }
    }
}
