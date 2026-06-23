using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetDefaultColumnWidthDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the default column width to 12 characters
                worksheet.Cells.StandardWidth = 12;

                // Save the workbook
                workbook.Save("DefaultColumnWidth.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetDefaultColumnWidthDemo.Run();
        }
    }
}