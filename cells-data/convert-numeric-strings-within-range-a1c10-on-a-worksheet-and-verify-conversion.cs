// Title: Convert numeric strings to numbers in cells A1:C10 and count numeric cells using Aspose.Cells for .NET (C#)
// AI Prompts: Scan the range A1:C10 with Aspose.Cells, identify string cells that represent numbers, and replace each with its numeric value. | After converting the strings, iterate the same range to tally how many cells are now of numeric type and output the total. | Save the modified workbook to a new file after the data‑type conversion.
// Common Searches: C# Aspose.Cells convert string cells to numbers in a specific range | How to count numeric cells after type conversion with Aspose.Cells | Parse numeric strings using invariant culture in Aspose.Cells C# example | Iterate over A1:C10 range and change data type in Aspose.Cells | Save workbook after converting string values to numeric in Aspose.Cells
// Tags: convert string cells to numeric Aspose.Cells | cell range iteration A1:C10 Aspose.Cells | numeric cell count after conversion C# | parse double invariant culture Aspose.Cells | save workbook after data type conversion Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;

// The example loads an Excel file, iterates through cells A1:C10, parses any string that can be interpreted as a double using invariant culture, replaces the string with the numeric value, counts how many cells in that range are now numeric, prints the count, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify by name)
        Worksheet sheet = workbook.Worksheets[0];

        // Define the target range A1:C10
        CellArea range = CellArea.CreateCellArea("A1", "C10");

        // Convert numeric strings to actual numbers
        for (int row = range.StartRow; row <= range.EndRow; row++)
        {
            for (int col = range.StartColumn; col <= range.EndColumn; col++)
            {
                Cell cell = sheet.Cells[row, col];

                // Process only string cells
                if (cell.Type == CellValueType.IsString)
                {
                    string text = cell.StringValue.Trim();

                    // Try parsing the string as a double using invariant culture
                    if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
                    {
                        // Replace the string with the numeric value
                        cell.PutValue(number);
                    }
                }
            }
        }

        // Verify conversion: count how many cells are now numeric
        int numericCount = 0;
        for (int row = range.StartRow; row <= range.EndRow; row++)
        {
            for (int col = range.StartColumn; col <= range.EndColumn; col++)
            {
                Cell cell = sheet.Cells[row, col];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    numericCount++;
                }
            }
        }

        Console.WriteLine($"Number of numeric cells after conversion: {numericCount}");

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
