using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel 97‑2003 workbook
            string filePath = "sample.xls";

            try
            {
                // Verify that the file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File \"{filePath}\" not found.");
                    return;
                }

                // Create LoadOptions for Excel 97‑2003 (XLS) format using the appropriate constructor
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Demonstrate that the workbook is loaded
                Console.WriteLine("Workbook loaded successfully.");
                Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine("An error occurred while loading the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}