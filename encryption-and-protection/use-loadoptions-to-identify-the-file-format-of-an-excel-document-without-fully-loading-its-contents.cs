// Title: Identify Excel file format using Aspose.Cells LoadOptions without loading the full workbook in C#
// AI Prompts: Show how to use Aspose.Cells LoadOptions to read only the file header and obtain the FileFormatType in C#. | Provide a C# snippet that determines whether an Excel file is XLS, XLSX, CSV, etc., using LoadOptions without loading worksheet data. | Generate code that safely checks an Excel file's format, includes file‑existence validation and exception handling with Aspose.Cells.
// Common Searches: how to get Excel file type with Aspose.Cells LoadOptions in C# | detect workbook format without opening full file Aspose.Cells | C# identify XLSX vs XLS using Aspose.Cells without loading data | Aspose.Cells loadoptions file format detection example
// Tags: Aspose.Cells LoadOptions format identification | detect Excel workbook type without full load | retrieve FileFormatType from Excel file C# | handle missing Excel file Aspose.Cells | exception handling for format detection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example validates the presence of an Excel file, uses Aspose.Cells LoadOptions to inspect only the file header, reads the Workbook.FileFormat property to obtain the detected FileFormatType, and prints the result while gracefully handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be inspected
        string filePath = "sample.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found at path '{filePath}'.");
            return;
        }

        try
        {
            // Load the workbook; Aspose.Cells automatically detects the format
            Workbook workbook = new Workbook(filePath);

            // Retrieve the detected format from the loaded workbook
            FileFormatType detectedFormat = workbook.FileFormat;

            // Display the identified file format
            Console.WriteLine($"Detected format: {detectedFormat}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
