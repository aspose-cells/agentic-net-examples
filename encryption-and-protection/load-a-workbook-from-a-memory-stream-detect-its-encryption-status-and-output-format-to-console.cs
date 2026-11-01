// Title: Load an Excel workbook from a MemoryStream and print its file format with Aspose.Cells for .NET
// Common Searches: how to read an Excel file from a byte array using Aspose.Cells in C# | retrieve the file format of a workbook loaded from a MemoryStream with Aspose.Cells | Aspose.Cells C# load workbook from stream without saving to disk | display workbook format type in console using Aspose.Cells .NET | check if an Excel workbook is encrypted when opening from a MemoryStream with Aspose.Cells
// Tags: memory stream workbook loading Aspose.Cells | retrieve workbook file format C# | Aspose.Cells read Excel from byte array | console display workbook format Aspose.Cells | encrypted workbook detection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example reads an Excel file into a byte array, wraps it in a MemoryStream, loads it into an Aspose.Cells Workbook, obtains the Workbook.FileFormat property, and writes the detected format to the console, including basic file existence validation and exception handling.
class Program
{
    static void Main()
    {
        const string filePath = "sample.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        try
        {
            // Load the workbook bytes into a memory stream
            byte[] workbookBytes = File.ReadAllBytes(filePath);
            using (MemoryStream memoryStream = new MemoryStream(workbookBytes))
            {
                // Load the workbook from the memory stream
                Workbook workbook = new Workbook(memoryStream);

                // Get the workbook's file format
                FileFormatType format = workbook.FileFormat;

                // Output the results to the console
                Console.WriteLine($"Format: {format}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
