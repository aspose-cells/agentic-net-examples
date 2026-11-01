// Title: Identify and log the file format of an unknown Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells detect unknown Excel workbook format programmatically | How to retrieve workbook file format type with Aspose.Cells .NET | Log the detected Excel file format after loading with Aspose.Cells | Determine if a file is XLS, XLSX, CSV, or ODS using Aspose.Cells in C# | Example of using Workbook.FileFormat property to identify Excel file type
// Tags: Aspose.Cells workbook file format detection | Workbook.FileFormat property C# | detect Excel file type Aspose.Cells | log workbook format .NET console | auto-detect spreadsheet format Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the file exists, loads it with Aspose.Cells (which auto‑detects the spreadsheet format), reads the Workbook.FileFormat property to get the identified format, and writes that format to the console while handling any loading errors.
class Program
{
    static void Main(string[] args)
    {
        // Path to the workbook whose format is unknown.
        // If a path is provided as a command‑line argument it will be used,
        // otherwise a default placeholder path is used.
        string filePath = args.Length > 0 ? args[0] : "unknownWorkbook";

        // Verify that the file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook. Aspose.Cells automatically detects the format.
            Workbook workbook = new Workbook(filePath);

            // Retrieve the detected file format.
            FileFormatType detectedFormat = workbook.FileFormat;

            // Log the identified format.
            Console.WriteLine($"Detected file format: {detectedFormat}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during loading or processing.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
