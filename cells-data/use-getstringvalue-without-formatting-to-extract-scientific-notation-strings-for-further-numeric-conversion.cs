// Title: Read raw scientific‑notation text from an Excel cell with GetStringValue (CellValueFormatStrategy.None) and convert to double – Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, writes the string "1.23E+05" to cell A1, extracts the exact cell content without any formatting using GetStringValue with CellValueFormatStrategy.None, and parses the scientific‑notation string to a double with invariant‑culture parsing.
// Keywords: Aspose.Cells GetStringValue | CellValueFormatStrategy.None | scientific notation Excel C# | unformatted cell text | convert string to double Aspose.Cells | read raw Excel value .NET | Excel scientific notation parsing
// Common Searches: GetStringValue raw text Aspose.Cells | retrieve scientific notation string from Excel cell C# | CellValueFormatStrategy.None example | convert Excel scientific notation to double | Aspose.Cells read unformatted cell value
// Developer Intent: Extract the exact scientific‑notation string from a worksheet cell without formatting and transform it into a numeric double.
// Use Cases: Importing data where numbers are stored as scientific‑notation strings and need precise numeric conversion. | Validating Excel exports that automatically display scientific notation but require the original text for calculations. | Cleaning and normalizing large spreadsheets by reading raw cell values before applying custom parsing logic.
// AI Prompts: Show how to use Aspose.Cells GetStringValue with CellValueFormatStrategy.None to read a scientific‑notation string from a cell and parse it to double in C#. | Provide a C# snippet that extracts unformatted text containing scientific notation from an Excel worksheet using Aspose.Cells and safely converts it to a numeric type. | Explain the effect of CellValueFormatStrategy.None on GetStringValue output and demonstrate its use for accurate numeric conversion of Excel scientific‑notation values.

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsScientificNotationDemo
{
    // This example creates a workbook, writes the string "1.23E+05" to cell A1, extracts the exact cell content without any formatting using GetStringValue with CellValueFormatStrategy.None, and parses the scientific‑notation string to a double with invariant‑culture parsing.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a scientific notation string into a cell
            // This simulates a value that appears in scientific format in Excel
            cells["A1"].PutValue("1.23E+05");

            // Retrieve the raw string without any formatting using GetStringValue with CellValueFormatStrategy.None
            string rawString = cells["A1"].GetStringValue(CellValueFormatStrategy.None);
            Console.WriteLine("Raw string from cell (no formatting): " + rawString);

            // Convert the scientific notation string to a numeric double
            if (double.TryParse(rawString, NumberStyles.Float | NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out double numericValue))
            {
                Console.WriteLine("Converted numeric value: " + numericValue);
            }
            else
            {
                Console.WriteLine("Failed to convert the string to a numeric value.");
            }

            // Save the workbook (optional, just to demonstrate lifecycle handling)
            workbook.Save("ScientificNotationDemo.xlsx");
        }
    }
}
