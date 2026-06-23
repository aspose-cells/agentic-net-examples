using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HideGridlinesAndFreezeHeader
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data to demonstrate the header row
                worksheet.Cells["A1"].PutValue("Header");
                worksheet.Cells["A2"].PutValue("Data 1");
                worksheet.Cells["A3"].PutValue("Data 2");
                worksheet.Cells["A4"].PutValue("Data 3");

                // Hide gridlines for a cleaner view
                worksheet.IsGridlinesVisible = false;

                // Freeze the first row (header) so it stays visible while scrolling
                worksheet.FreezePanes(1, 0, 1, 0);

                // Save the workbook
                string outputPath = "HideGridlinesAndFreezeHeader.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            HideGridlinesAndFreezeHeader.Run();
        }
    }
}