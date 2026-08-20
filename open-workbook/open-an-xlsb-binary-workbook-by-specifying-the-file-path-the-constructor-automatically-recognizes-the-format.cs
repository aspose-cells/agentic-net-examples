// Title: Open XLSB Workbook from File Path with Aspose.Cells for .NET
// Description: Demonstrates loading a binary XLSB workbook by passing its full path to the Aspose.Cells Workbook constructor, which auto‑detects the format. Includes file‑existence validation, basic workbook metadata output, reading cell A1, and robust exception handling.
// Keywords: Aspose.Cells open XLSB | load binary Excel workbook C# | Workbook constructor format detection | check file existence Aspose.Cells | read cell value XLSB Aspose | C# .NET Excel binary file | exception handling Aspose.Cells
// Common Searches: how to open xlsb with aspose.cells c# | aspose.cells automatically detect workbook format | read first worksheet name from xlsb file | validate xlsb file exists before loading | c# example loading binary excel workbook
// Developer Intent: Load an XLSB file from a specified path, verify its presence, and access worksheet and cell data using Aspose.Cells.
// Use Cases: Open an XLSB workbook and display the total worksheet count and the name of the first sheet. | Read the value of cell A1 from the first worksheet of an XLSB file. | Ensure the XLSB file exists before creating a Workbook instance to avoid runtime errors. | Handle exceptions gracefully when loading a binary Excel workbook.
// AI Prompts: Generate C# code that opens an XLSB workbook with Aspose.Cells, checks file existence, and prints worksheet details. | Show how to catch and log exceptions while loading a binary Excel file using Aspose.Cells. | Provide an example that reads multiple cell values from the first worksheet of an XLSB workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates loading a binary XLSB workbook by passing its full path to the Aspose.Cells Workbook constructor, which auto‑detects the format. Includes file‑existence validation, basic workbook metadata output, reading cell A1, and robust exception handling.
    public class OpenXlsbWorkbookDemo
    {
        public static void Run()
        {
            // Specify the full path to the XLSB file.
            string xlsbPath = @"C:\Data\sample.xlsb";

            // Verify that the file exists before attempting to load it.
            if (!File.Exists(xlsbPath))
            {
                Console.WriteLine($"Error: The file \"{xlsbPath}\" was not found.");
                return;
            }

            try
            {
                // The Workbook constructor automatically detects the file format (XLSB in this case).
                Workbook workbook = new Workbook(xlsbPath);

                // Access the first worksheet to verify that the file was loaded successfully.
                Worksheet sheet = workbook.Worksheets[0];

                // Output some basic information about the loaded workbook.
                Console.WriteLine($"Workbook loaded from: {xlsbPath}");
                Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
                Console.WriteLine($"First worksheet name: {sheet.Name}");

                // Example: read the value of cell A1 if it exists.
                if (sheet.Cells["A1"].Value != null)
                {
                    Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");
                }
                else
                {
                    Console.WriteLine("Cell A1 is empty.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application.
    public class Program
    {
        public static void Main(string[] args)
        {
            OpenXlsbWorkbookDemo.Run();
        }
    }
}
