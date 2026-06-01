using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetColumnWidthPixelDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the Cells collection
                Cells cells = worksheet.Cells;

                // Set the width of column 2 (third column, zero‑based index) to 150 pixels
                cells.SetColumnWidthPixel(2, 150);

                // Optionally put some data to visualize the column width
                cells["C1"].PutValue("Column width set to 150 pixels");

                // Define output file path
                string outputPath = "SetColumnWidthPixelDemo.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            SetColumnWidthPixelDemo.Run();
        }
    }
}