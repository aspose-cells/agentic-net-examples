using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MaxDisplayRangeDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["A2"].PutValue(100);
            cells["B2"].PutValue(200);
            cells["A3"].PutValue(300);
            cells["B3"].PutValue(400);

            Aspose.Cells.Range maxDisplayRange = cells.MaxDisplayRange;

            if (maxDisplayRange != null)
            {
                Console.WriteLine("Max Display Range:");
                Console.WriteLine($"Start Row (zero‑based): {maxDisplayRange.FirstRow}");
                Console.WriteLine($"Start Column (zero‑based): {maxDisplayRange.FirstColumn}");
                Console.WriteLine($"Total Rows: {maxDisplayRange.RowCount}");
                Console.WriteLine($"Total Columns: {maxDisplayRange.ColumnCount}");
            }
            else
            {
                Console.WriteLine("Worksheet is empty; MaxDisplayRange is null.");
            }

            workbook.Save("MaxDisplayRangeDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            MaxDisplayRangeDemo.Run();
        }
    }
}