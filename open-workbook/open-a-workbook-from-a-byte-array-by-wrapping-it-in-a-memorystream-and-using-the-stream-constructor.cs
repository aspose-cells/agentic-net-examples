// Title: Open an Excel workbook from a byte[] using Aspose.Cells .NET (MemoryStream)
// Description: Demonstrates how to read an XLSX file into a byte array, wrap it in a MemoryStream, instantiate a Workbook from the stream, read the first worksheet name, and save the workbook to a new file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells load workbook from byte array | MemoryStream Excel .NET | Workbook constructor stream | read Excel bytes C# | save workbook after stream load
// Common Searches: open Excel file from byte[] Aspose.Cells | Aspose.Cells MemoryStream example | load workbook from byte array C# | read first worksheet name from byte array Aspose
// Developer Intent: Load an Excel workbook from a byte array via MemoryStream and optionally save or inspect it.
// Use Cases: Process an uploaded Excel file received as a byte[] in a web service without writing to disk. | Convert an in‑memory Excel byte array to another format (e.g., PDF, CSV) by opening it with a MemoryStream. | Validate the worksheet name of a byte‑array Excel payload before further business logic.
// AI Prompts: Provide C# code that opens an Excel workbook from a byte[] with Aspose.Cells, changes cell B2 on the second sheet, and returns the modified file as a byte array. | Generate a method that accepts a byte[] Excel file, reads the first worksheet name using Aspose.Cells, and logs it. | Explain best practices for handling large Excel files when loading them from a byte array with Aspose.Cells to minimize memory consumption.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Demonstrates how to read an XLSX file into a byte array, wrap it in a MemoryStream, instantiate a Workbook from the stream, read the first worksheet name, and save the workbook to a new file with Aspose.Cells for .NET.
    public class OpenWorkbookFromByteArrayDemo
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load Excel file bytes
                byte[] excelData = File.ReadAllBytes(inputPath);

                // Wrap the byte array in a MemoryStream
                using (MemoryStream stream = new MemoryStream(excelData))
                {
                    // Open the workbook from the stream
                    Workbook workbook = new Workbook(stream);

                    // Access the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];
                    Console.WriteLine($"First worksheet name: {sheet.Name}");

                    // Save the workbook to a new file
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
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
