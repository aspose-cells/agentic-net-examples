using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Byte array containing an Excel file (replace with actual data as needed)
        byte[] excelBytes = GetSampleExcelBytes();

        // Wrap the byte array in a MemoryStream
        using (MemoryStream stream = new MemoryStream(excelBytes))
        {
            // Open the workbook from the stream using the Workbook(Stream) constructor
            Workbook workbook = new Workbook(stream);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: read the value from cell A1
            Console.WriteLine($"A1 value: {worksheet.Cells["A1"].StringValue}");

            // Save the workbook to a file (optional)
            workbook.Save("WorkbookFromBytes.xlsx");
        }
    }

    // Helper method to generate a sample Excel file and return its bytes
    static byte[] GetSampleExcelBytes()
    {
        using (MemoryStream ms = new MemoryStream())
        {
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].Value = "Sample data";
            wb.Save(ms, SaveFormat.Xlsx);
            return ms.ToArray();
        }
    }
}

// Author: Aspose.Cells .NET example