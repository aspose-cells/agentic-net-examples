// Title: How to loop through all worksheets in an Excel file using Aspose.Cells for .NET and log each sheet’s name and index
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, iterates over workbook.Worksheets, and prints each worksheet’s Name and zero‑based Index to the console. | Show a robust example that verifies the input file exists, loads the workbook, and uses a foreach loop to output worksheet identifiers for diagnostic logging. | Provide a method that returns a List<int> containing all worksheet indexes after enumerating the Worksheets collection with Aspose.Cells.
// Common Searches: Aspose.Cells .NET get worksheet index and name while iterating | C# enumerate worksheets in a workbook and log sheet IDs using Aspose.Cells | How to perform diagnostic logging of sheet identifiers in Aspose.Cells | Loop through workbook.Worksheets collection Aspose.Cells example | Retrieve zero‑based worksheet indexes in Aspose.Cells for .NET
// Tags: Aspose.Cells iterate worksheets .NET | Aspose.Cells get worksheet index | Aspose.Cells log sheet names and IDs | Aspose.Cells load workbook and enumerate sheets | Aspose.Cells diagnostic worksheet iteration

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel workbook with Aspose.Cells, safely checks the file, iterates through each worksheet, and writes the worksheet name and its zero‑based index to the console for diagnostic reporting.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook from the specified file
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            try
            {
                // Iterate over all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Use the worksheet's Index as its identifier (zero‑based)
                    int sheetIndex = sheet.Index;

                    // Diagnostic reporting
                    Console.WriteLine($"Worksheet Name: {sheet.Name}, Index: {sheetIndex}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while processing worksheets: {ex.Message}");
                return;
            }

            // (Optional) Save the workbook if any modifications were made
            // Uncomment the following lines if you need to save changes
            /*
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
            */
        }
    }
}
