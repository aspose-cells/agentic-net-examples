// Title: Convert string‑formatted numbers to numeric values only in the first worksheet using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook, iterate the used range of the first worksheet, and replace any cell that holds a numeric string with a double value via Aspose.Cells. | Process only the first sheet of a workbook, detect string cells that can be parsed as numbers, convert them to numeric type, and save the file while leaving all other worksheets untouched.
// Common Searches: Aspose.Cells C# convert numeric text to numbers in first worksheet only | How to change string numbers to numeric values on the first sheet with Aspose.Cells .NET | Iterate used range of first worksheet and replace text numbers with doubles in C#
// Tags: convert string to numeric Aspose.Cells | first worksheet only processing | replace string cell with double | iterate used range cells Aspose.Cells | preserve other worksheets Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The example loads an Excel file, scans the used range of the first worksheet, converts cells containing numeric text strings to double values, saves the workbook, and ensures all other worksheets remain unchanged.
class ConvertStringNumbersInFirstSheet
{
    static void Main()
    {
        // Load the workbook (replace with your actual input file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet firstSheet = workbook.Worksheets[0];
        Cells cells = firstSheet.Cells;

        // Determine the used range to limit iteration
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through each cell in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // Process only cells that currently contain a string
                if (cell.Type == CellValueType.IsString)
                {
                    string text = cell.StringValue;

                    // Try to parse the string as a double using invariant culture
                    if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
                    {
                        // Replace the string with the numeric value
                        cell.PutValue(numericValue);
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output file path)
        workbook.Save("output.xlsx");
    }
}
