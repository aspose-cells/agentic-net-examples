using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CellsEnumeratorXlsxDemo
    {
        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("John");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Alice");
            worksheet.Cells["B3"].PutValue(25);
            worksheet.Cells["C4"].PutValue(DateTime.Now);

            // Get the cells enumerator for the worksheet
            IEnumerator cellEnumerator = worksheet.Cells.GetEnumerator();

            // Iterate through all instantiated cells
            Console.WriteLine("Enumerating cells in the worksheet:");
            while (cellEnumerator.MoveNext())
            {
                // The enumerator returns objects of type Cell
                Cell cell = (Cell)cellEnumerator.Current;

                // Output cell address and its value (if any)
                if (cell.Value != null)
                {
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }
            }

            // Save the workbook as an XLSX file
            workbook.Save("CellsEnumeratorDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CellsEnumeratorXlsxDemo.Run();
        }
    }
}