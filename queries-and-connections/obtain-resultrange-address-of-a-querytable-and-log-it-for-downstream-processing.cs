using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class QueryTableResultRangeLogger
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Mary");

            try
            {
                if (worksheet.QueryTables.Count > 0)
                {
                    QueryTable queryTable = worksheet.QueryTables[0];
                    Aspose.Cells.Range resultRange = queryTable.ResultRange;

                    Console.WriteLine("QueryTable ResultRange Address: " + resultRange.Address);
                    Console.WriteLine("Row Count: " + resultRange.RowCount);
                    Console.WriteLine("Column Count: " + resultRange.ColumnCount);
                }
                else
                {
                    Console.WriteLine("No query tables found in the worksheet.");
                }

                workbook.Save("QueryTableResultRangeDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            QueryTableResultRangeLogger.Run();
        }
    }
}