using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsQueryTableResultRange
{
    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            if (worksheet.QueryTables.Count > 0)
            {
                QueryTable queryTable = worksheet.QueryTables[0];
                AsposeRange resultRange = queryTable.ResultRange;

                Console.WriteLine("ResultRange Address: " + resultRange.Address);
                Console.WriteLine("First Row: " + resultRange.FirstRow);
                Console.WriteLine("First Column: " + resultRange.FirstColumn);
                Console.WriteLine("Row Count: " + resultRange.RowCount);
                Console.WriteLine("Column Count: " + resultRange.ColumnCount);
            }
            else
            {
                Console.WriteLine("No query tables found in the first worksheet.");
            }

            workbook.Save("output.xlsx");
        }
    }
}