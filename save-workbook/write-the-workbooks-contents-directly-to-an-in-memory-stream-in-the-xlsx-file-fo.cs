using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (uses the Workbook() constructor rule)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(100);

        // Save the workbook directly to an in‑memory stream in XLSX format
        using (MemoryStream stream = new MemoryStream())
        {
            // Uses the Save(Stream, SaveFormat) method rule
            workbook.Save(stream, SaveFormat.Xlsx);

            // Reset the stream position if further reading is required
            stream.Position = 0;

            // Example output: display the size of the generated XLSX content
            Console.WriteLine($"Workbook saved to stream. Size: {stream.Length} bytes");
        }

        // Clean up resources
        workbook.Dispose();
    }
}