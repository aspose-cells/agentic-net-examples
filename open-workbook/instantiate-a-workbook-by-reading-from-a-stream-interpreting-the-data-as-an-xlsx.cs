using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a sample workbook and write it to a memory stream in XLSX format
        using (MemoryStream sourceStream = new MemoryStream())
        {
            Workbook sampleWorkbook = new Workbook(); // default empty workbook
            sampleWorkbook.Worksheets[0].Cells["A1"].PutValue("Hello from stream");
            sampleWorkbook.Save(sourceStream, SaveFormat.Xlsx);
            sourceStream.Position = 0; // reset stream for reading

            // Load a workbook from the stream using the Workbook(Stream) constructor
            Workbook workbook = new Workbook(sourceStream);

            // Verify that the data was loaded correctly
            Console.WriteLine(workbook.Worksheets[0].Cells["A1"].StringValue);

            // Save the loaded workbook to a file (optional)
            workbook.Save("LoadedFromStream.xlsx", SaveFormat.Xlsx);
        }
    }
}