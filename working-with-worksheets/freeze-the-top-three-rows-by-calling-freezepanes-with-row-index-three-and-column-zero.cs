using System;
using Aspose.Cells;
using System.IO;

namespace AsposeCellsExamples
{
    public class FreezeTopThreeRowsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Freeze the top three rows (row index 3, column index 0)
                // freezedRows = 3 (rows to freeze), freezedColumns = 0 (no columns frozen)
                worksheet.FreezePanes(3, 0, 3, 0);

                // Define output file path
                string outputPath = "FreezeTopThreeRows.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            FreezeTopThreeRowsDemo.Run();
        }
    }
}