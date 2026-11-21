// Title: How to load a macro-enabled XLSM workbook from a file path into an Aspose.Cells Workbook in C#
// Common Searches: c# code to open a macro-enabled xlsm file with Aspose.Cells | aspose.cells load xlsm workbook from specific folder in C# | check if Excel file exists before loading with Aspose.Cells C# example | handle errors when opening a macro-enabled workbook using Aspose.Cells
// Tags: load XLSM workbook Aspose.Cells C# | Aspose.Cells workbook constructor file path | verify Excel file existence before loading | exception handling Aspose.Cells workbook load

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample checks that the target XLSM file exists, then creates a new Aspose.Cells Workbook by passing the file path to the constructor, and reports success or any loading errors via console output.
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the full path to the existing XLSM workbook
            string filePath = @"C:\Path\To\YourWorkbook.xlsm";

            // Prevent FileNotFoundException by checking file existence
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the file into a Workbook object
                Workbook workbook = new Workbook(filePath);
                Console.WriteLine("Workbook loaded successfully.");
                // Further processing can be performed on 'workbook' here
            }
            catch (Exception ex)
            {
                // Handle any runtime errors during loading
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }
}
