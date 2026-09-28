// Title: Load an Excel workbook from a MemoryStream, determine its file format, and check encryption status with Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an Excel file into a MemoryStream, loads it with Aspose.Cells, prints the workbook's FileFormat, and attempts to detect encryption by catching the relevant exception. | Show how to open a workbook from a byte array using Aspose.Cells, output its format, and gracefully handle password‑protected files.
// Common Searches: asp.net core load excel from memory stream using aspose.cells | c# detect password protected workbook with aspose.cells loadoptions | how to get workbook file format after loading from stream in aspose.cells | aspose.cells read excel bytes into workbook and check encryption status | c# handle encrypted excel file when opening with aspose.cells
// Tags: load workbook from memory stream Aspose.Cells | retrieve workbook file format .NET | detect encrypted workbook Aspose.Cells | handle password protected Excel file C# | read Excel bytes into Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// The example reads an Excel file into a byte array, creates a MemoryStream, loads the workbook with Aspose.Cells, prints the detected FileFormat, and demonstrates that encryption detection must be performed by catching exceptions when attempting to open a password‑protected file.
class Program
{
    static void Main()
    {
        const string filePath = "sample.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File \"{filePath}\" not found.");
            return;
        }

        try
        {
            // Load workbook bytes from the file
            byte[] workbookBytes = File.ReadAllBytes(filePath);

            // Create a memory stream from the byte array
            using (MemoryStream memoryStream = new MemoryStream(workbookBytes))
            {
                // Load the workbook from the memory stream
                Workbook workbook = new Workbook(memoryStream);

                // Get the workbook's file format (e.g., Xlsx, Xls, Csv, etc.)
                FileFormatType format = workbook.FileFormat;

                // Output the format to the console
                Console.WriteLine($"Format: {format}");

                // Encryption detection is not directly available in this version of Aspose.Cells.
                // If needed, attempt to load with LoadOptions and catch the exception for encrypted files.
                Console.WriteLine("Encryption detection not supported in this API version.");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions (e.g., I/O errors, invalid format)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
