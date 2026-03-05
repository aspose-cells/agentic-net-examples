using System;
using Aspose.Cells;

namespace AsposeCellsNumericConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate cells with numeric values stored as strings
            cells[0, 0].PutValue("123");          // integer as string
            cells[0, 1].PutValue("45.67");        // decimal as string
            cells[0, 2].PutValue("2021-06-20");   // date string (will be converted to a date number)
            cells[0, 3].PutValue("NotANumber");   // non‑numeric string, remains unchanged

            // Convert all possible string values in the worksheet to their numeric equivalents
            cells.ConvertStringToNumericValue();

            // Verify conversion (optional)
            Console.WriteLine("A1 (numeric): " + cells[0, 0].DoubleValue);
            Console.WriteLine("B1 (numeric): " + cells[0, 1].DoubleValue);
            Console.WriteLine("C1 (date as OADate): " + cells[0, 2].DoubleValue);
            Console.WriteLine("D1 (string): " + cells[0, 3].StringValue);

            // Save the workbook to XLSX format
            workbook.Save("ConvertedData.xlsx", SaveFormat.Xlsx);
        }
    }
}