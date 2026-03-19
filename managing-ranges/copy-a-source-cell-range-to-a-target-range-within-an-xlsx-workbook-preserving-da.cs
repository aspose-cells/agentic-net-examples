using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source range A1:E5 with sample data and formatting
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    // Put a string value
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");

                    // Apply a simple style to demonstrate formatting copy
                    Style style = workbook.CreateStyle();
                    style.Font.Color = Color.Blue;
                    style.Font.IsBold = (row + col) % 2 == 0;
                    sheet.Cells[row, col].SetStyle(style);
                }
            }

            // Define source and destination ranges (both 5 rows x 5 columns)
            // Source: A1:E5 (row 0, col 0, 5 rows, 5 columns)
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange(0, 0, 5, 5);
            // Destination: G1:K5 (row 0, col 6, 5 rows, 5 columns)
            Aspose.Cells.Range destinationRange = sheet.Cells.CreateRange(0, 6, 5, 5);

            // Copy source range to destination range preserving data and formatting
            destinationRange.Copy(sourceRange);

            // Save the workbook to an XLSX file
            workbook.Save("RangeCopyResult.xlsx");
        }
    }
}