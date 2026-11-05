// Title: Detect the format of an unknown Excel workbook and log it using Aspose.Cells for .NET
// Common Searches: how to automatically detect Excel file format with Aspose.Cells in C# | using LoadOptions with LoadFormat.Auto to identify workbook type in .NET | retrieve workbook.FileFormat after loading unknown spreadsheet with Aspose.Cells | C# code to log detected spreadsheet format when opening unknown file | handle missing file and default workbook format using Aspose.Cells
// Tags: Aspose.Cells LoadOptions automatic detection | Workbook.FileFormat retrieval | identify unknown Excel workbook .NET | log spreadsheet format C# | missing file handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks whether a specified file exists; if not, it creates an empty Workbook and prints its default format. When the file is present, it loads the workbook with LoadOptions set to LoadFormat.Auto, reads the detected format via the Workbook.FileFormat property, and writes the result to the console, with exception handling for load errors.
    class Program
    {
        static void Main()
        {
            // Path to the workbook file (replace with a valid path if available)
            string filePath = "unknownWorkbook";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                // Optionally create an empty workbook to demonstrate default format handling
                Workbook emptyWorkbook = new Workbook();
                Console.WriteLine($"Created an empty workbook. Default format: {emptyWorkbook.FileFormat}");
                return;
            }

            try
            {
                // Load the workbook with automatic format detection
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Retrieve and display the detected file format
                FileFormatType detectedFormat = workbook.FileFormat;
                Console.WriteLine($"Detected workbook format: {detectedFormat}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }
}
