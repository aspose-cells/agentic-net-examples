using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSM workbook
            string filePath = @"C:\Data\SampleWorkbook.xlsm";

            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Show the name of the first worksheet (if any)
                if (workbook.Worksheets.Count > 0)
                {
                    Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
                }
                else
                {
                    Console.WriteLine("The workbook contains no worksheets.");
                }
            }
            catch (Exception ex)
            {
                // Handle unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}