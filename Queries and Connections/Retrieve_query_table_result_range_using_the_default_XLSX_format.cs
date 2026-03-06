using System;
using Aspose.Cells;

namespace AsposeCellsQueryTableResultRange
{
    class Program
    {
        static void Main(string[] args)
        {
            Workbook workbook;
            try
            {
                workbook = new Workbook("InputWithQueryTable.xlsx");
            }
            catch
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];

                ws.Cells["A1"].PutValue("ID");
                ws.Cells["B1"].PutValue("Name");
                ws.Cells["A2"].PutValue(1);
                ws.Cells["B2"].PutValue("John");
                ws.Cells["A3"].PutValue(2);
                ws.Cells["B3"].PutValue("Mary");
            }

            Worksheet worksheet = workbook.Worksheets[0];

            if (worksheet.QueryTables.Count > 0)
            {
                QueryTable queryTable = worksheet.QueryTables[0];
                Aspose.Cells.Range resultRange = queryTable.ResultRange;

                Console.WriteLine("ResultRange Address: " + resultRange.Address);
                Console.WriteLine("Rows: " + resultRange.RowCount + ", Columns: " + resultRange.ColumnCount);
                Console.WriteLine("First Row Index: " + resultRange.FirstRow);
                Console.WriteLine("First Column Index: " + resultRange.FirstColumn);

                foreach (Cell cell in resultRange)
                {
                    Console.WriteLine($"Cell[{cell.Row}, {cell.Column}] = {cell.Value}");
                }
            }
            else
            {
                Console.WriteLine("No query tables found in the worksheet.");
            }

            workbook.Save("QueryTableResultRangeDemo.xlsx");
            Console.WriteLine("Workbook saved as QueryTableResultRangeDemo.xlsx");
        }
    }
}