// Title: C# – Convert Numeric Strings to Numbers in Range A1:C10 with Aspose.Cells and Verify Conversion
// Description: This Aspose.Cells for .NET example creates a workbook, fills cells A1:C10 with mixed string data, uses Cell.PutValue with the conversion flag to turn numeric‑looking strings into true numbers, checks each cell with IsNumericValue to report the result, and saves the file as ConvertedRange.xlsx.
// Keywords: Aspose.Cells | C# | .NET | convert string to number | PutValue conversion flag | IsNumericValue | numeric string conversion | worksheet range conversion | A1:C10 | data cleaning | CSV import to Excel | Excel automation example | GitHub code sample
// Common Searches: Aspose.Cells convert numeric strings in a range | C# PutValue conversion flag example | How to check if a cell is numeric Aspose.Cells | Convert text numbers to numbers in Excel using Aspose.Cells | Verify numeric conversion in Aspose.Cells .NET
// Developer Intent: Transform every string cell that represents a number inside A1:C10 into a numeric type and confirm the conversion programmatically.
// Use Cases: Clean up CSV imports where numbers are stored as text before calculations. | Prepare data for charting or pivot tables by ensuring numeric values are real numbers. | Automate report generation that requires numeric formatting after converting string entries.
// AI Prompts: Generate a C# snippet that uses Aspose.Cells to convert numeric strings to numbers in a specified range and list the cells that were changed. | Explain the behavior of Cell.PutValue with the conversion flag and how to use IsNumericValue to validate the conversion. | Suggest enhancements to handle locale‑specific formats such as commas, spaces, or different decimal separators during string‑to‑number conversion with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    // This Aspose.Cells for .NET example creates a workbook, fills cells A1:C10 with mixed string data, uses Cell.PutValue with the conversion flag to turn numeric‑looking strings into true numbers, checks each cell with IsNumericValue to report the result, and saves the file as ConvertedRange.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in the range A1:C10
            // Some cells contain numeric strings, others contain non‑numeric strings
            string[,] sampleData = new string[10, 3]
            {
                { "123",   "45.67",   "ABC" },
                { "00100", "0.001",   "2021-01-01" },
                { "-50",   "3.14",    "Hello" },
                { "7e2",   "NaN",     "World" },
                { "0",     "1000",    "Test" },
                { "12.34", "56.78",   "Sample" },
                { "9",     "8",       "7" },
                { "1.2.3", "4,5",     "6" },
                { "   9",  "10 ",     "Eleven" },
                { "12",    "13.0",    "14" }
            };

            for (int row = 0; row < 10; row++)
                for (int col = 0; col < 3; col++)
                    cells[row, col].PutValue(sampleData[row, col]);

            // Convert numeric strings within the range A1:C10
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    Cell cell = cells[row, col];
                    // Attempt conversion only if the cell currently holds a string
                    if (cell.Type == CellValueType.IsString)
                    {
                        // PutValue with conversion flag will try to convert the string to a numeric type
                        cell.PutValue(cell.StringValue, true, false);
                    }
                }
            }

            // Verify conversion and output results
            Console.WriteLine("Verification of conversion in range A1:C10:");
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    Cell cell = cells[row, col];
                    string address = cell.Name;
                    if (cell.IsNumericValue)
                    {
                        Console.WriteLine($"{address}: Converted to numeric, Value = {cell.Value}");
                    }
                    else
                    {
                        Console.WriteLine($"{address}: Remains non‑numeric, Value = \"{cell.StringValue}\"");
                    }
                }
            }

            // Save the workbook
            workbook.Save("ConvertedRange.xlsx");
        }
    }
}
