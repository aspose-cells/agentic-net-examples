using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Save the workbook to a memory stream in XLSX format using default options
        using (MemoryStream stream = new MemoryStream())
        {
            // Use the Save(Stream, SaveFormat) overload as defined in the API
            workbook.Save(stream, SaveFormat.Xlsx);

            // Reset the stream position if further processing is required
            stream.Position = 0;

            // Optional: write the stream to a physical file for verification
            using (FileStream file = new FileStream("output.xlsx", FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(file);
            }
        }

        // Clean up resources
        workbook.Dispose();
    }
}