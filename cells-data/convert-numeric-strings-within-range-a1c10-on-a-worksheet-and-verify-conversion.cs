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

            // Populate the range A1:C10 with a mix of numeric strings, dates, and non‑numeric strings
            string[,] sampleData = new string[10, 3]
            {
                { "123",   "45.67",   "2021-06-20" },
                { "abc",   "0",       "100"       },
                { "3.1415","-200",    "text"      },
                { "00100", "5e2",     "12/31/2020"},
                { "7",     "8.9",     "invalid"   },
                { "0",     "NaN",     "123abc"    },
                { "42",    "3.0",     "2022/01/01"},
                { "99",    "100.00",  "hello"     },
                { "5",     "6",       "7"         },
                { "10",    "20",      "30"        }
            };

            // Write the sample data into the worksheet
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue(sampleData[row, col]);
                }
            }

            // Convert all string values that can be interpreted as numeric (or date) to their native types
            cells.ConvertStringToNumericValue();

            // Verify conversion by iterating over the same range and checking IsNumericValue
            Console.WriteLine("Verification of conversion in range A1:C10:");
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    Cell cell = cells[row, col];
                    string address = cell.Name; // e.g., "A1"
                    if (cell.IsNumericValue)
                    {
                        // Numeric values can be accessed via DoubleValue (covers int, double, date)
                        Console.WriteLine($"{address}: Numeric ({cell.DoubleValue})");
                    }
                    else
                    {
                        Console.WriteLine($"{address}: Non‑numeric (\"{cell.StringValue}\")");
                    }
                }
            }

            // Save the workbook to verify the result visually if needed
            workbook.Save("ConvertedRange.xlsx");
        }
    }
}