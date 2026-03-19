using System;
using Aspose.Cells;

namespace AsposeCellsLoadXlsxExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Load the workbook using the string constructor (loads the file)
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Display some basic information about the worksheet
            Console.WriteLine("Worksheet Name: " + worksheet.Name);
            Console.WriteLine("Number of Cells: " + worksheet.Cells.Count);

            // Optionally, read a cell value (e.g., A1) if it exists
            if (worksheet.Cells["A1"].Value != null)
            {
                Console.WriteLine("Cell A1 Value: " + worksheet.Cells["A1"].Value.ToString());
            }

            // Save the workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("Workbook loaded from '" + sourcePath + "' and saved as '" + outputPath + "'.");
        }
    }
}