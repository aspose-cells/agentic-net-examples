// Title: Load an Excel workbook from a MemoryStream, edit cells, and save back to a MemoryStream using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook from MemoryStream byte array | read cell value after loading XLSX from stream with Aspose.Cells | modify worksheet cells and save to another MemoryStream using Aspose.Cells | using LoadOptions to specify Xlsx format when loading from stream Aspose.Cells | how to create and save an in‑memory Excel file with Aspose.Cells .NET
// Tags: memory stream workbook loading Aspose.Cells | loadoptions xlsx stream initialization | cell A1 read Aspose.Cells | cell B2 write Aspose.Cells | memory stream workbook saving Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a simple XLSX workbook in memory, loads it from a MemoryStream with LoadOptions, reads the value of cell A1, writes a new value to cell B2, and then saves the modified workbook to a separate MemoryStream.
class Program
{
    static void Main()
    {
        try
        {
            // Create a sample workbook in memory and obtain its byte array.
            byte[] excelData = CreateSampleWorkbook();

            // Load the workbook from the memory stream.
            using (MemoryStream inputStream = new MemoryStream(excelData))
            {
                // Optional: specify load options (e.g., format).
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Initialize the Workbook with the stream and load options.
                Workbook workbook = new Workbook(inputStream, loadOptions);

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Read and display the value of cell A1.
                Cell cellA1 = worksheet.Cells["A1"];
                Console.WriteLine($"Cell A1 value: {cellA1.Value}");

                // Demonstrate a simple modification: write a new value to cell B2.
                worksheet.Cells["B2"].PutValue("Added via stream");

                // Save the modified workbook to another memory stream.
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    Console.WriteLine($"Modified workbook saved to memory stream ({outputStream.Length} bytes).");
                }
            }
        }
        catch (Exception ex)
        {
            // Report any errors without suppressing them.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to create a simple workbook and return its bytes.
    static byte[] CreateSampleWorkbook()
    {
        using (MemoryStream ms = new MemoryStream())
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Hello from memory");
            wb.Save(ms, SaveFormat.Xlsx);
            return ms.ToArray();
        }
    }
}
