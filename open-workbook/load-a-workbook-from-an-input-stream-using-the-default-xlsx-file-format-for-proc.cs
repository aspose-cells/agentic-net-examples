using System;
using System.IO;
using Aspose.Cells;

class LoadWorkbookFromStream
{
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Create a sample workbook and write it to a memory stream.
        // This simulates receiving an XLSX file via an input stream.
        using (MemoryStream inputStream = new MemoryStream())
        {
            using (Workbook sample = new Workbook())
            {
                sample.Worksheets[0].Cells["A1"].PutValue("Hello from stream");
                sample.Save(inputStream, SaveFormat.Xlsx);
            }

            // Reset the stream position so it can be read from the beginning.
            inputStream.Position = 0;

            // Load the workbook from the stream using the Workbook(Stream) constructor.
            Workbook workbook = new Workbook(inputStream);

            // Example processing: read the value from cell A1.
            string loadedValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Loaded cell A1: " + loadedValue);

            // Save the workbook to a file for verification.
            workbook.Save("LoadedFromStream.xlsx", SaveFormat.Xlsx);
        }
    }
}