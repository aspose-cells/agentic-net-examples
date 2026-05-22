using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class AdjustColumnWidthWithStandardWidthDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Set the default column width (base width) before adding any data
                // Width is specified in character units
                cells.StandardWidth = 20.0;

                // Populate some sample data to demonstrate that the base width is applied
                cells["A1"].PutValue("Short");
                cells["B1"].PutValue("This is a longer piece of text");
                cells["C1"].PutValue("Another column with medium length");

                // Verify the standard width and actual width of a column
                Console.WriteLine("Standard Width set to: " + cells.StandardWidth);
                Console.WriteLine("Column 0 actual width: " + cells.GetColumnWidth(0));

                // Define output file path
                string outputPath = "AdjustColumnWidthWithStandardWidthDemo.xlsx";

                // Save the workbook (lifecycle: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public static class Program
    {
        public static void Main(string[] args)
        {
            AdjustColumnWidthWithStandardWidthDemo.Run();
        }
    }
}