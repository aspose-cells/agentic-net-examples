using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file
            string excelPath = @"C:\Data\Sample.xlsx";

            try
            {
                // Ensure the file exists; create a minimal workbook if it does not
                if (!File.Exists(excelPath))
                {
                    // Create directory structure if missing
                    string? dir = Path.GetDirectoryName(excelPath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    // Create a new workbook with a default worksheet and save it
                    Workbook newWb = new Workbook();
                    newWb.Worksheets[0].Name = "Sheet1";
                    newWb.Save(excelPath);
                }

                // Load the workbook from the existing file
                Workbook workbook = new Workbook(excelPath);

                // Access the first worksheet to verify that the file was loaded
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"Loaded workbook: {excelPath}");
                Console.WriteLine($"First worksheet name: {sheet.Name}");
                Console.WriteLine($"Number of cells in first worksheet: {sheet.Cells.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}