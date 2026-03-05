using System;
using Aspose.Cells;

namespace AsposeCellsEnumeratorDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);
            sheet.Cells["A4"].PutValue("Charlie");
            sheet.Cells["B4"].PutValue(78);

            // 1. Enumerate all cells in the worksheet
            Console.WriteLine("Enumerating all cells:");
            foreach (Cell cell in sheet.Cells)
            {
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // 2. Enumerate cells in the first row (header row)
            Console.WriteLine("\nEnumerating cells in first row:");
            foreach (Cell cell in sheet.Cells.Rows[0])
            {
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // 3. Enumerate rows in the worksheet
            Console.WriteLine("\nEnumerating rows:");
            foreach (Row row in sheet.Cells.Rows)
            {
                Cell firstCell = sheet.Cells[row.Index, 0];
                string firstValue = firstCell?.Value?.ToString() ?? "Empty";
                Console.WriteLine($"Row {row.Index}: {firstValue}");
            }

            // 4. Enumerate cells in a specific range (B2:C4)
            Console.WriteLine("\nEnumerating cells in range B2:C4:");
            Aspose.Cells.Range range = sheet.Cells.CreateRange("B2:C4");
            foreach (Cell cell in range)
            {
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // Save the workbook as XLSX
            workbook.Save("EnumeratorDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}