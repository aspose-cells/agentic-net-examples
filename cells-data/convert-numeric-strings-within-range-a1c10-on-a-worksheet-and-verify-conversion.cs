// Title: Convert numeric strings to numbers in range A1:C10 using Aspose.Cells for .NET and verify each cell's type
// AI Prompts: Write C# code with Aspose.Cells that populates cells A1:C10 with alternating numeric‑string and text values, invokes ConvertStringToNumericValue, and logs whether each cell was converted to a numeric type. | Show how to loop through a defined range after calling ConvertStringToNumericValue to output the cell address, its resulting data type, and the numeric value when conversion succeeds.
// Common Searches: Aspose.Cells C# convert string values to numbers in a specific range and check conversion result | How to use ConvertStringToNumericValue on cells A1:C10 and determine which cells became numeric | C# Aspose.Cells example for populating mixed data and converting numeric strings to numeric types | Verify cell data type after ConvertStringToNumericValue in Aspose.Cells .NET | Convert numeric string to double in Excel worksheet using Aspose.Cells API
// Tags: Aspose.Cells ConvertStringToNumericValue example | numeric string to double conversion Aspose.Cells | C# iterate over range A1:C10 cells | verify cell type after conversion Aspose.Cells | populate worksheet with mixed string data Aspose.Cells

using System;
using Aspose.Cells;

// Creates a workbook, fills A1:C10 with alternating numeric‑string and text entries, runs ConvertStringToNumericValue to turn convertible strings into numeric values, iterates the range to output each cell's address, resulting data type and value, and saves the file as ConvertedRange.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate range A1:C10 with mixed data (numeric strings and regular text)
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                // Even sum of indices -> numeric string, odd sum -> non‑numeric text
                if ((row + col) % 2 == 0)
                {
                    // Example numeric string: "1.1", "2.2", etc.
                    string numericString = (row + 1).ToString() + "." + (col + 1).ToString();
                    cells[row, col].PutValue(numericString);
                }
                else
                {
                    // Example non‑numeric string
                    cells[row, col].PutValue($"Text{row + 1}{col + 1}");
                }
            }
        }

        // Convert all string data in the worksheet to numeric values where possible
        // (this will affect the numeric strings in A1:C10)
        cells.ConvertStringToNumericValue();

        // Verify conversion for each cell in the range A1:C10
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                Cell cell = cells[row, col];
                string address = cell.Name; // e.g., "A1"

                if (cell.IsNumericValue)
                {
                    // Cell was successfully converted to a numeric type
                    Console.WriteLine($"{address} converted to numeric: {cell.DoubleValue}");
                }
                else
                {
                    // Cell remains a string (non‑numeric)
                    Console.WriteLine($"{address} remains string: {cell.StringValue}");
                }
            }
        }

        // Save the workbook to verify the result in Excel
        workbook.Save("ConvertedRange.xlsx");
    }
}
