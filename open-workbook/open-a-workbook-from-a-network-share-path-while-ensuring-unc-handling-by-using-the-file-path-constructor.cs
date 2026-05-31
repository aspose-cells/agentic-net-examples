using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkShareDemo
{
    class Program
    {
        static void Main()
        {
            // UNC path to the Excel file on a network share
            string uncPath = @"\\ServerName\ShareFolder\example.xlsx";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(uncPath))
            {
                Console.WriteLine($"File not found: {uncPath}");
                return;
            }

            try
            {
                // Load the workbook from the UNC path
                Workbook workbook = new Workbook(uncPath);

                // Display the name of the first worksheet
                Worksheet firstSheet = workbook.Worksheets[0];
                Console.WriteLine("First worksheet name: " + firstSheet.Name);

                // Display the absolute path of the opened workbook
                Console.WriteLine("Workbook absolute path: " + workbook.AbsolutePath);
            }
            catch (Exception ex)
            {
                // Handle any runtime errors (e.g., access issues, corrupted file)
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}