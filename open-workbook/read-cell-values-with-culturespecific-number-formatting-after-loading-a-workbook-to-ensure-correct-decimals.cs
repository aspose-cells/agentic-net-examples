// Title: Read Excel cell numeric value with French locale using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read cell value with French number format | How to set workbook culture to fr-FR in Aspose.Cells .NET | Retrieve double value from Excel cell respecting locale in Aspose.Cells | Parse Excel cell string to double using CultureInfo with Aspose.Cells | Load workbook and read localized numeric value in C# Aspose.Cells
// Tags: Aspose.Cells set workbook CultureInfo | read numeric cell value Aspose.Cells | parse cell string with CultureInfo C# | Excel locale-specific number parsing Aspose | Workbook.Settings.CultureInfo usage

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, assigns the French (fr-FR) CultureInfo to the workbook's Settings, accesses cell A1 on the first worksheet, and obtains its numeric value both via the DoubleValue property (which respects the workbook's culture) and by manually parsing the cell's raw string with the same CultureInfo. It also demonstrates handling of password‑protected files and generic exception handling.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (no password). If the file is password‑protected,
                // a CellsException will be thrown.
                var loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Set the culture that matches the number format in the workbook
                workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

                // Access the first worksheet and cell A1
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];

                // Option 1: Use DoubleValue – respects the workbook's CultureInfo setting
                double valueFromDouble = cell.DoubleValue;

                // Option 2: Read the raw string and parse it manually with the same CultureInfo
                string rawString = cell.StringValue;
                double valueFromString = double.Parse(rawString, workbook.Settings.CultureInfo);

                // Output the results
                Console.WriteLine($"Value via DoubleValue: {valueFromDouble}");
                Console.WriteLine($"Value via parsed string: {valueFromString}");
            }
            catch (CellsException ex)
            {
                // Detect password protection via message content (fallback for API versions without ErrorCode)
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("The workbook is password protected. Provide a password in LoadOptions.");
                }
                else
                {
                    Console.WriteLine($"CellsException: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
