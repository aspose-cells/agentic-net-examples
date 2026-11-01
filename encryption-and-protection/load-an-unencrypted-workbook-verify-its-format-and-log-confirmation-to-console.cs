// Title: Load an unencrypted Excel workbook, detect its format, and log the result using Aspose.Cells for .NET
// Common Searches: C# load Excel file with Aspose.Cells and get workbook file format | How to verify an Excel workbook is XLSX using Aspose.Cells in .NET | Aspose.Cells detect workbook format and handle password‑protected file exception | Read Excel workbook with Aspose.Cells and output format to console in C# | Check if Excel file exists and load it without decryption using Aspose.Cells
// Tags: load workbook Aspose.Cells C# | detect Excel file format Aspose.Cells | verify Xlsx format Aspose.Cells | catch password protection exception Aspose.Cells | output workbook load result to console

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the file's existence, loads an unencrypted workbook with Aspose.Cells, retrieves the detected FileFormat, prints the format and a confirmation of XLSX status to the console, and catches exceptions such as password protection or corruption.
class Program
{
    static void Main()
    {
        // Path to the workbook file (replace with your actual file path)
        string filePath = "input.xlsx";

        // Prevent FileNotFoundException by checking file existence first
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook (unencrypted). If the file is password‑protected,
            // Aspose.Cells will throw an exception which we catch below.
            Workbook workbook = new Workbook(filePath);

            // Get and display the detected file format
            FileFormatType format = workbook.FileFormat;
            Console.WriteLine($"Workbook loaded successfully. Detected format: {format}");

            // Optional check for a specific format
            if (format == FileFormatType.Xlsx)
            {
                Console.WriteLine("The workbook is in the expected XLSX format.");
            }
            else
            {
                Console.WriteLine("The workbook is not in XLSX format.");
            }
        }
        catch (Exception ex)
        {
            // Handle errors such as password‑protected files or corrupted data
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
