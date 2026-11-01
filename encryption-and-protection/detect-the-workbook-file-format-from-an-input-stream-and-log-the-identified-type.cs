// Title: Detect Excel workbook format from a stream and log the result using Aspose.Cells for .NET
// Common Searches: how to determine Excel file format from a stream with Aspose.Cells in C# | using LoadOptions Unknown to detect workbook type in Aspose.Cells | retrieve Workbook.FileFormat after loading a workbook stream | log detected workbook format in a C# console app using Aspose.Cells | detect Excel workbook format when file extension is missing
// Tags: Aspose.Cells loadoptions unknown format detection | C# stream-based workbook format detection | auto-detect Excel format with Aspose.Cells | console logging of detected workbook format | Workbook.FileFormat property retrieval

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the Excel file exists, opens it as a read‑only stream, creates a LoadOptions instance with LoadFormat.Unknown so Aspose.Cells auto‑detects the workbook format, loads the workbook, reads the Workbook.FileFormat property, and writes the detected FileFormatType to the console while handling exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Open the workbook stream; Aspose.Cells will auto‑detect the format.
            using (Stream inputStream = File.OpenRead(inputPath))
            {
                var loadOptions = new LoadOptions(LoadFormat.Unknown);
                var workbook = new Workbook(inputStream, loadOptions);

                // Get the detected file format.
                FileFormatType detectedFormat = workbook.FileFormat;

                // Log the identified format.
                Console.WriteLine($"Detected workbook format: {detectedFormat}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
