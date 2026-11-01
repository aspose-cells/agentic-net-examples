// Title: Add built‑in and custom document properties to an Excel workbook loaded from a MemoryStream and save it back to a MemoryStream with Aspose.Cells for .NET
// Common Searches: how to set built‑in Excel properties from a byte array using Aspose.Cells in C# | Aspose.Cells load workbook from MemoryStream and modify metadata without writing to disk | add custom document properties to an in‑memory Excel file with Aspose.Cells .NET | save updated workbook to a byte array using Aspose.Cells SaveFormat.Xlsx | update Title, Author, and Company properties of a workbook loaded from stream in C#
// Tags: memory stream workbook loading Aspose.Cells | built‑in document properties update Aspose.Cells | custom document properties addition .NET | save workbook to byte array Aspose.Cells | in‑memory Excel manipulation .NET

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an existing Excel workbook from a byte array via a MemoryStream, setting built‑in properties (Title, Author, Company) and adding custom properties (ProjectId, Reviewed) with Aspose.Cells, then saving the modified workbook into another MemoryStream as Xlsx and retrieving the updated byte array—all without accessing the file system.
class Program
{
    static void Main()
    {
        try
        {
            // Assume we have an existing workbook as a byte array (could be from a database, network, etc.)
            byte[] existingWorkbookBytes = GetWorkbookBytes();

            // Load the workbook from a memory stream
            using (MemoryStream loadStream = new MemoryStream(existingWorkbookBytes))
            {
                // Create a Workbook instance from the memory stream
                Workbook workbook = new Workbook(loadStream);

                // ---- Add or modify metadata (document properties) ----
                // Built‑in properties (use property setters, not indexer)
                workbook.BuiltInDocumentProperties.Title = "Sales Report Q4";
                workbook.BuiltInDocumentProperties.Author = "John Doe";
                workbook.BuiltInDocumentProperties.Company = "Acme Corp";

                // Custom properties (optional)
                workbook.CustomDocumentProperties.Add("ProjectId", 12345);
                workbook.CustomDocumentProperties.Add("Reviewed", true);

                // ---- Save the modified workbook back to a memory stream ----
                using (MemoryStream saveStream = new MemoryStream())
                {
                    // Save the workbook into the memory stream (in the same format as original)
                    workbook.Save(saveStream, SaveFormat.Xlsx);

                    // At this point, saveStream contains the updated workbook bytes.
                    // If you need the byte array, call ToArray().
                    byte[] updatedWorkbookBytes = saveStream.ToArray();

                    // Example usage: send the bytes over a network, store in a database, etc.
                    ProcessUpdatedWorkbook(updatedWorkbookBytes);
                }
            }
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Placeholder method to obtain the original workbook bytes.
    static byte[] GetWorkbookBytes()
    {
        // In a real scenario, replace this with actual retrieval logic.
        // For demonstration, create a simple workbook and return its bytes.
        using (MemoryStream ms = new MemoryStream())
        {
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Hello, Aspose!");
            wb.Save(ms, SaveFormat.Xlsx);
            return ms.ToArray();
        }
    }

    // Placeholder method to handle the updated workbook bytes.
    static void ProcessUpdatedWorkbook(byte[] data)
    {
        // Example: write to console the size of the updated workbook.
        Console.WriteLine($"Updated workbook size: {data.Length} bytes");
    }
}
