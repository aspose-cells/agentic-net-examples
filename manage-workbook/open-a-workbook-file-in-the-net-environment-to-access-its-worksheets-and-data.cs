using System;
using Aspose.Cells;

namespace AsposeCellsOpenWorkbookDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file to be opened
            string filePath = "input.xlsx";

            // Open the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(filePath);

            // Access the collection of worksheets
            WorksheetCollection sheets = workbook.Worksheets;

            // Iterate through each worksheet and display its name and the value of cell A1 (if any)
            for (int i = 0; i < sheets.Count; i++)
            {
                Worksheet sheet = sheets[i];
                Console.WriteLine($"Worksheet {i}: {sheet.Name}");

                // Retrieve the value from cell A1
                var cellValue = sheet.Cells["A1"].Value;
                if (cellValue != null)
                {
                    Console.WriteLine($"  A1 = {cellValue}");
                }
                else
                {
                    Console.WriteLine("  A1 is empty");
                }
            }

            // Optionally, save the workbook after any modifications (none in this example)
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}