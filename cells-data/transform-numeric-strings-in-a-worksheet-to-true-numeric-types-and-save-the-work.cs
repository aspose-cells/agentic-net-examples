using System;
using Aspose.Cells;

namespace NumericStringConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet's cells
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate cells with numeric strings and a non‑numeric string
            cells[0, 0].PutValue("123.45");   // numeric string
            cells[0, 1].PutValue("67");       // numeric string
            cells[0, 2].PutValue("2021-06-20"); // date string (will become numeric date)
            cells[0, 3].PutValue("NotANumber"); // remains a string

            // Convert all possible string values to their native numeric types
            cells.ConvertStringToNumericValue();

            // Optional: display the resulting types for verification
            Console.WriteLine($"A1 Type: {cells[0, 0].Type}, Value: {cells[0, 0].Value}");
            Console.WriteLine($"B1 Type: {cells[0, 1].Type}, Value: {cells[0, 1].Value}");
            Console.WriteLine($"C1 Type: {cells[0, 2].Type}, Value: {cells[0, 2].Value}");
            Console.WriteLine($"D1 Type: {cells[0, 3].Type}, Value: {cells[0, 3].Value}");

            // Save the workbook as an XLSX file (lifecycle: save)
            workbook.Save("NumericStringsConverted.xlsx", SaveFormat.Xlsx);
        }
    }
}