using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ParseNumericStringsToXlsx
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate cells with string values, some of which represent numbers or dates
            cells[0, 0].PutValue("123");          // integer string
            cells[0, 1].PutValue("45.67");        // floating‑point string
            cells[0, 2].PutValue("2023-08-15");   // date string (ISO format)
            cells[0, 3].PutValue("NotANumber");   // non‑numeric string
            cells[0, 4].PutValue("9.99E2");       // scientific notation string
            cells[0, 5].PutValue("12/31/2021");   // US date format string

            // Convert all convertible string values to their proper numeric or date types
            cells.ConvertStringToNumericValue();

            // Demonstrate the conversion results in the console
            Console.WriteLine("Conversion results:");
            for (int col = 0; col <= 5; col++)
            {
                Cell cell = cells[0, col];
                Console.WriteLine($"Cell {cell.Name}: Value = {cell.Value}, Type = {cell.Type}");
            }

            // Save the workbook to an XLSX file
            string outputPath = "ParsedNumericStrings.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ParseNumericStringsToXlsx.Run();
        }
    }
}