using System;
using Aspose.Cells;

namespace AsposeCellsNumbersStoredAsTextDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate cells with values that are stored as text
            cells[0, 0].PutValue("123");          // numeric text
            cells[0, 1].PutValue("45.67");        // numeric text with decimal
            cells[0, 2].PutValue("2021-06-20");   // date text
            cells[0, 3].PutValue("NotANumber");   // non‑numeric text

            // Convert all string values that can be interpreted as numbers or dates
            cells.ConvertStringToNumericValue();

            // Display the converted values to verify the conversion
            Console.WriteLine("A1 (numeric): " + cells[0, 0].DoubleValue);          // 123
            Console.WriteLine("B1 (numeric): " + cells[0, 1].DoubleValue);          // 45.67
            Console.WriteLine("C1 (date): " + cells[0, 2].DateTimeValue.ToShortDateString()); // 6/20/2021
            Console.WriteLine("D1 (string): " + cells[0, 3].StringValue);          // NotANumber

            // Save the workbook to an Excel file
            workbook.Save("NumbersStoredAsTextConverted.xlsx");
        }
    }
}