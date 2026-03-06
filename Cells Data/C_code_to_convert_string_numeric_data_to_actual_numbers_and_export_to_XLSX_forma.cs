using System;
using Aspose.Cells;

namespace AsposeCellsStringToNumberExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate cells with string representations of numbers and dates
            cells[0, 0].PutValue("123");          // integer as string
            cells[0, 1].PutValue("45.67");        // decimal as string
            cells[0, 2].PutValue("2021-06-20");   // date as string
            cells[0, 3].PutValue("NotANumber");   // non‑numeric string

            // Convert all possible string values to their native numeric/date types
            // Uses Cells.ConvertStringToNumericValue method
            cells.ConvertStringToNumericValue();

            // Optional: display the converted values to verify
            Console.WriteLine("A1 numeric: " + cells[0, 0].DoubleValue);
            Console.WriteLine("B1 numeric: " + cells[0, 1].DoubleValue);
            Console.WriteLine("C1 date   : " + cells[0, 2].DateTimeValue);
            Console.WriteLine("D1 string : " + cells[0, 3].StringValue);

            // Save the workbook as XLSX (lifecycle save)
            workbook.Save("StringToNumberOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}