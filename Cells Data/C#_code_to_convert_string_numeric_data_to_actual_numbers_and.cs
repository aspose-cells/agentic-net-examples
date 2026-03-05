using System;
using Aspose.Cells;

namespace AsposeCellsStringToNumberExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate cells with string representations of numbers and dates
            cells[0, 0].PutValue("123.45");          // numeric string
            cells[0, 1].PutValue("678");            // integer string
            cells[0, 2].PutValue("2023-08-15");     // date string
            cells[0, 3].PutValue("NotANumber");     // non‑numeric string

            // Convert all possible string values to their native numeric/date types
            cells.ConvertStringToNumericValue();

            // Verify conversion (optional)
            Console.WriteLine("A1 (double): " + cells[0, 0].DoubleValue);
            Console.WriteLine("B1 (double): " + cells[0, 1].DoubleValue);
            Console.WriteLine("C1 (date): " + cells[0, 2].DateTimeValue);
            Console.WriteLine("D1 (string): " + cells[0, 3].StringValue);

            // Save the workbook as XLSX
            workbook.Save("StringToNumberExport.xlsx", SaveFormat.Xlsx);
        }
    }
}