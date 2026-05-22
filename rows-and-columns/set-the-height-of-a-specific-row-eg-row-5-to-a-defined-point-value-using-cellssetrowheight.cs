using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetRowHeightDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the height of row 5 (zero‑based index) to 30 points
                worksheet.Cells.SetRowHeight(5, 30.0);

                // Define output file path
                string outputPath = "RowHeightDemo.xlsx";

                // Save the workbook
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
            SetRowHeightDemo.Run();
        }
    }
}