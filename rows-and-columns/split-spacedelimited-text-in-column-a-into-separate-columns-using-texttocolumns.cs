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

            // Sample space‑delimited data in column A
            cells["A1"].PutValue("John Doe 30");
            cells["A2"].PutValue("Jane Smith 28");
            cells["A3"].PutValue("Bob Johnson 45");

            // Configure load options to use space as the delimiter
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.Separator = ' ';                     // space character
            loadOptions.TreatConsecutiveDelimitersAsOne = true; // optional: ignore multiple spaces

            // Split the text in column A (column index 0) for the first 3 rows
            // Parameters: start row, start column, total rows, options
            cells.TextToColumns(0, 0, 3, loadOptions);

            // Verify the split (optional)
            Console.WriteLine("After TextToColumns:");
            Console.WriteLine($"B1: {cells["B1"].StringValue}"); // Expected: Doe
            Console.WriteLine($"C1: {cells["C1"].StringValue}"); // Expected: 30
            Console.WriteLine($"B2: {cells["B2"].StringValue}"); // Expected: Smith
            Console.WriteLine($"C2: {cells["C2"].StringValue}"); // Expected: 28

            // Save the workbook
            workbook.Save("SplitResult.xlsx");
        }
    }
}