using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class GridlinesScreenPrintDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure gridlines are visible on screen
                worksheet.IsGridlinesVisible = true;

                // Hide gridlines when the sheet is printed
                worksheet.PageSetup.PrintGridlines = false;

                // Add some sample data to visualize the gridlines on screen
                worksheet.Cells["A1"].PutValue("Screen Gridlines Visible");
                worksheet.Cells["A2"].PutValue("Printed Gridlines Hidden");

                // Define output file path
                string outputPath = "GridlinesScreenPrintDemo.xlsx";

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
            GridlinesScreenPrintDemo.Run();
        }
    }
}