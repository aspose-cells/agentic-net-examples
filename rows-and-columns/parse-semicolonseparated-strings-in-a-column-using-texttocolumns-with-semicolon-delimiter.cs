using System;
using Aspose.Cells;

namespace AsposeCellsTextToColumnsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with semicolon‑separated strings
            cells["A1"].PutValue("Apple;Red;1.2");
            cells["A2"].PutValue("Banana;Yellow;0.8");
            cells["A3"].PutValue("Cherry;Red;0.2");

            // Configure TextToColumns options to use semicolon as the delimiter
            TxtLoadOptions options = new TxtLoadOptions();
            options.Separator = ';';               // character delimiter
            // options.SeparatorString = ";";      // alternatively, string delimiter

            // Apply TextToColumns starting from row 0, column 0 (A1)
            // totalRows = number of rows to process (3 in this example)
            int totalColumnsCreated = cells.TextToColumns(0, 0, 3, options);

            // Output the result to the console for verification
            Console.WriteLine($"Columns created: {totalColumnsCreated}");
            for (int row = 0; row < 3; row++)
            {
                Console.WriteLine(
                    $"Row {row + 1}: " +
                    $"{cells[row, 0].StringValue}, " +   // Fruit name
                    $"{cells[row, 1].StringValue}, " +   // Color
                    $"{cells[row, 2].StringValue}");    // Weight
            }

            // Save the workbook (optional)
            workbook.Save("SemicolonSplitResult.xlsx");
        }
    }
}