using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);

        // Configure JSON save options.
        // JSON output from Aspose.Cells is UTF-8 by default, so no explicit encoding property exists.
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            Indent = "  " // optional pretty‑print indentation
        };

        // Save the workbook as JSON using a memory stream.
        // The library writes UTF‑8 bytes to the stream.
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, jsonOptions);
            // Write the UTF‑8 bytes to a file.
            File.WriteAllBytes("output.json", stream.ToArray());
        }
    }
}