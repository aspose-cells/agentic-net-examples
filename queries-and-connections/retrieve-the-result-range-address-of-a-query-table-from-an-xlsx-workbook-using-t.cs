using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsQueryTableResultRange
{
    class Program
    {
        static void Main(string[] args)
        {
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            if (worksheet.QueryTables.Count > 0)
            {
                QueryTable queryTable = worksheet.QueryTables[0];
                AsposeRange resultRange = queryTable.ResultRange;
                Console.WriteLine("ResultRange Address: " + resultRange.Address);
            }
            else
            {
                Console.WriteLine("No query tables found in the worksheet.");
            }

            workbook.Save("output.xlsx");
        }
    }
}