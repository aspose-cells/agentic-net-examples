using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate range A1:C10 with mixed data (numeric strings, dates, non‑numeric strings)
            string[,] sampleData = new string[10, 3]
            {
                { "100",   "200.5",   "Hello" },
                { "300",   "04/01/2022", "World" },
                { "400.75","500",     "123ABC" },
                { "600",   "700",     "800" },
                { "900",   "ABC",     "1.23E3" },
                { "2021-12-31", "0", "Text" },
                { "123",   "456",     "789" },
                { "1.5",   "2.5",     "3.5" },
                { "1000",  "2000",    "3000" },
                { "NotNumber", "Another", "100.00" }
            };

            // Fill the cells with the sample data
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    // PutValue(string) stores the value as a string initially
                    cells[row, col].PutValue(sampleData[row, col]);
                }
            }

            // Convert all string data in the worksheet to numeric values where possible
            // This uses the Cells.ConvertStringToNumericValue method as required
            cells.ConvertStringToNumericValue();

            // Verify conversion for each cell in the range A1:C10
            Console.WriteLine("Verification of conversion in range A1:C10:");
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    Cell cell = cells[row, col];
                    string address = cell.Name; // e.g., "A1"

                    if (cell.IsNumericValue)
                    {
                        // Numeric value detected; display as double
                        Console.WriteLine($"{address}: Numeric = {cell.DoubleValue}");
                    }
                    else
                    {
                        // Still a string; display the original string value
                        Console.WriteLine($"{address}: String  = \"{cell.StringValue}\"");
                    }
                }
            }

            // Save the workbook to verify the result visually if needed
            workbook.Save("ConvertedRange.xlsx");
        }
    }
}