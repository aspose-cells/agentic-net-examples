using System;
using Aspose.Cells;

namespace AsposeCellsNumberConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate cells with numeric values stored as text
            cells[0, 0].PutValue("123");          // A1
            cells[0, 1].PutValue("45.67");        // B1
            cells[0, 2].PutValue("9,876");        // C1 (will be treated as text)
            cells[0, 3].PutValue("NotANumber");   // D1 (remains as string)

            // Convert all string data that can be interpreted as numbers to numeric values
            // Rule used: Cells.ConvertStringToNumericValue()
            cells.ConvertStringToNumericValue();

            // Verify conversion (optional)
            Console.WriteLine("A1 numeric value: " + cells[0, 0].DoubleValue);
            Console.WriteLine("B1 numeric value: " + cells[0, 1].DoubleValue);
            Console.WriteLine("C1 type after conversion: " + cells[0, 2].Type); // May remain string if format not recognized
            Console.WriteLine("D1 remains string: " + cells[0, 3].StringValue);

            // Save the workbook as XLSX (lifecycle: save)
            workbook.Save("ConvertedNumbers.xlsx", SaveFormat.Xlsx);
        }
    }
}