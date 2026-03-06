using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamDemo
{
    class Program
    {
        static void Main()
        {
            // Create a workbook and put some data in the first worksheet
            Workbook originalWorkbook = new Workbook();
            Worksheet sheet = originalWorkbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello from stream!");

            // Save the workbook to a memory stream in the default XLSX format
            using (MemoryStream stream = new MemoryStream())
            {
                originalWorkbook.Save(stream, SaveFormat.Xlsx);
                // Reset the stream position to the beginning for reading
                stream.Position = 0;

                // Open a new workbook from the same stream using the Stream constructor
                Workbook loadedWorkbook = new Workbook(stream);

                // Access the first worksheet and read the value we wrote earlier
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                string cellValue = loadedSheet.Cells["A1"].StringValue;

                Console.WriteLine($"Loaded cell A1 value: {cellValue}");

                // Optionally, save the loaded workbook to a file to verify the operation
                loadedWorkbook.Save("LoadedFromStream.xlsx");
            }

            // Clean up
            originalWorkbook.Dispose();
        }
    }
}