using System;
using Aspose.Cells;

namespace AsposeCellsQueryTableUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be examined.
            // You can change this to any valid .xlsx/.xls file path.
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check if the worksheet contains any query tables.
                if (sheet.QueryTables.Count > 0)
                {
                    // Output the name of the worksheet that has query tables.
                    Console.WriteLine($"Worksheet with QueryTables: {sheet.Name}");
                }
            }

            // Optionally, save the workbook (no changes made in this utility).
            // This demonstrates the required save lifecycle rule.
            workbook.Save("output.xlsx");
        }
    }
}