using System;
using Aspose.Cells;

namespace AsposeCellsLoadSample
{
    class Program
    {
        static void Main()
        {
            // Path to the XLSX file that needs to be loaded.
            // Replace this with the actual path of your runtime sample file.
            string filePath = "sample.xlsx";

            // Load the workbook using the constructor that accepts a file path.
            // This utilizes the Workbook(string) overload as defined in the Aspose.Cells API.
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet to verify that the workbook was loaded successfully.
            Worksheet sheet = workbook.Worksheets[0];

            // Output some basic information about the loaded workbook.
            Console.WriteLine($"Workbook loaded from: {filePath}");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
            Console.WriteLine($"First worksheet name: {sheet.Name}");

            // Optionally, display the value of cell A1 if it exists.
            if (sheet.Cells["A1"].Value != null)
            {
                Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].Value}");
            }
            else
            {
                Console.WriteLine("Cell A1 is empty.");
            }

            // No saving is performed here as the task only requires loading the workbook.
        }
    }
}