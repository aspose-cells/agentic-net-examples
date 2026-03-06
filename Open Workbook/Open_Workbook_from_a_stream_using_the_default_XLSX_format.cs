using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamDemo
{
    class Program
    {
        static void Main()
        {
            // Create a workbook and put some data in the first cell
            using (Workbook sourceWorkbook = new Workbook())
            {
                sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Hello from stream");

                // Save the workbook to a memory stream in the default XLSX format
                using (MemoryStream stream = new MemoryStream())
                {
                    sourceWorkbook.Save(stream, SaveFormat.Xlsx);
                    stream.Position = 0; // Reset the stream position for reading

                    // Load a new workbook from the memory stream (default format is XLSX)
                    Workbook loadedWorkbook = new Workbook(stream);

                    // Read the value from the loaded workbook to verify it was opened correctly
                    string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
                    Console.WriteLine($"Loaded cell A1 value: {cellValue}");

                    // Save the loaded workbook to a physical file
                    loadedWorkbook.Save("LoadedFromStream.xlsx");
                }
            }
        }
    }
}