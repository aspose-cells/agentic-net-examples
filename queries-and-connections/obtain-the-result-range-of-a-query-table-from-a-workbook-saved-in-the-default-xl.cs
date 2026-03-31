using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsQueryTableResultRange
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (default XLSX format)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through worksheets to find query tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of query tables in the worksheet
                QueryTableCollection queryTables = sheet.QueryTables;

                if (queryTables.Count > 0)
                {
                    // Get the first query table (or iterate as needed)
                    QueryTable queryTable = queryTables[0];

                    // Obtain the result range of the query table
                    AsposeRange resultRange = queryTable.ResultRange;

                    // Display information about the result range
                    Console.WriteLine($"Worksheet: {sheet.Name}");
                    Console.WriteLine($"Query Table Name: {queryTable.Name}");
                    Console.WriteLine($"Result Range Address: {resultRange.Address}");
                    Console.WriteLine($"First Row: {resultRange.FirstRow}, First Column: {resultRange.FirstColumn}");
                    Console.WriteLine($"Row Count: {resultRange.RowCount}, Column Count: {resultRange.ColumnCount}");

                    // Optionally, iterate through cells in the result range
                    foreach (Cell cell in resultRange)
                    {
                        Console.WriteLine($"Cell [{cell.Row}, {cell.Column}] = {cell.Value}");
                    }
                }
                else
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' contains no query tables.");
                }
            }

            // Save the workbook (if any changes were made)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}