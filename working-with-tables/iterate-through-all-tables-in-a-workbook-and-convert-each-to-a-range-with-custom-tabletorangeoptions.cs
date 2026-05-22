using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableToRange
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of tables (ListObjects) on the current worksheet
                ListObjectCollection tables = sheet.ListObjects;

                // Iterate backwards because ConvertToRange removes the table from the collection
                for (int i = tables.Count - 1; i >= 0; i--)
                {
                    ListObject table = tables[i];

                    // Create custom options for conversion
                    TableToRangeOptions options = new TableToRangeOptions();

                    // Example: set the last row to the current end row of the table.
                    // This demonstrates using a custom option; adjust as needed.
                    options.LastRow = table.EndRow;

                    // Convert the table to a normal range using the options
                    table.ConvertToRange(options);
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}