using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class OpenWorkbookFromByteArrayDemo
    {
        public static void Run()
        {
            try
            {
                // Obtain a byte array that contains an Excel file.
                byte[] excelBytes = GetSampleExcelBytes();

                // Wrap the byte array in a MemoryStream.
                using (MemoryStream memoryStream = new MemoryStream(excelBytes))
                {
                    // Ensure the stream is positioned at the beginning.
                    memoryStream.Position = 0;

                    // Load the workbook from the stream.
                    Workbook workbook = new Workbook(memoryStream);

                    // Access the first worksheet and read a cell value.
                    Worksheet sheet = workbook.Worksheets[0];
                    Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

                    // Save the loaded workbook to a new file.
                    string outputPath = "LoadedFromBytes.xlsx";
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        private static byte[] GetSampleExcelBytes()
        {
            // Create a simple workbook in memory and return its byte representation.
            using (MemoryStream tempStream = new MemoryStream())
            {
                Workbook tempWorkbook = new Workbook();
                tempWorkbook.Worksheets[0].Cells["A1"].PutValue("Hello from byte array");
                tempWorkbook.Save(tempStream, SaveFormat.Xlsx);
                return tempStream.ToArray();
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            OpenWorkbookFromByteArrayDemo.Run();
        }
    }
}