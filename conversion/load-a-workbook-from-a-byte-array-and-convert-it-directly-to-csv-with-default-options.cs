using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvConversion
{
    class Program
    {
        static void Main()
        {
            // Example byte array containing an Excel file.
            // In a real scenario replace this with the actual byte[] data.
            byte[] excelBytes = File.ReadAllBytes("input.xlsx");

            // Load the workbook from the byte array using a MemoryStream.
            using (MemoryStream stream = new MemoryStream(excelBytes))
            {
                // Workbook(Stream) constructor loads the workbook from the stream.
                Workbook workbook = new Workbook(stream);

                // Save the workbook directly as CSV using default options.
                // Save(string, SaveFormat) is the provided rule for saving.
                workbook.Save("output.csv", SaveFormat.Csv);
            }

            Console.WriteLine("Conversion to CSV completed.");
        }
    }
}