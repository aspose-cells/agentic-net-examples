// Title: Load a macro‑enabled XLSM workbook from a file path into an Aspose.Cells Workbook using C#
// AI Prompts: Generate C# code that verifies a .xlsm file exists and creates an Aspose.Cells Workbook from the given path with proper exception handling. | Show how to instantiate an Aspose.Cells Workbook for a macro‑enabled Excel file and output a success or error message. | Provide a C# example that opens an existing .xlsm workbook with Aspose.Cells, validates the file, and logs the loading result.
// Common Searches: c# aspocells open existing macro enabled workbook from disk | how to read .xlsm file with Aspose.Cells in .NET | load xlsm workbook using Aspose.Cells C# example with file existence check | Aspose.Cells Workbook constructor path parameter for macro enabled Excel | exception handling when loading .xlsm with Aspose.Cells C#
// Tags: xlsm workbook loading Aspose.Cells C# | macro-enabled Excel import Aspose.Cells | Workbook constructor with file path Aspose.Cells | pre‑load file existence check C# | catch exceptions during Aspose.Cells workbook creation

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The sample checks that a specified .xlsm file exists, then creates an Aspose.Cells Workbook from that path, printing a success message or catching and displaying any loading exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSM workbook (modify as needed)
            string filePath = @"C:\Path\To\YourWorkbook.xlsm";

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(filePath);
                Console.WriteLine("Workbook loaded successfully.");
                // Additional processing can be performed here
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }
}
