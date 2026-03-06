using System;
using System.Data;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook("Introduction.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            int lastRow = cells.MaxDataRow;
            int lastColumn = cells.MaxDataColumn;

            int totalRows = lastRow + 1;
            int totalColumns = lastColumn + 1;

            AsposeRange usedRange = cells.CreateRange(0, 0, totalRows, totalColumns);
            usedRange.Name = "IntroUsedRange";

            DataTable dt = usedRange.ExportDataTable();

            Console.WriteLine($"Exported DataTable has {dt.Rows.Count} rows and {dt.Columns.Count} columns.");
            Console.WriteLine("First few rows:");

            int rowsToShow = Math.Min(5, dt.Rows.Count);
            for (int i = 0; i < rowsToShow; i++)
            {
                DataRow row = dt.Rows[i];
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    Console.Write($"{row[j]}\t");
                }
                Console.WriteLine();
            }

            workbook.Save("Introduction_Reviewed.xlsx");
        }
    }
}