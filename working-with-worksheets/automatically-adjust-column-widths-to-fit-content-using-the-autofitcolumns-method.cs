using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class AutoFitColumnsExample
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate cells with varying length text
                worksheet.Cells["A1"].PutValue("This is a test string");
                worksheet.Cells["B1"].PutValue("Another longer test string for demonstration");
                worksheet.Cells["C1"].PutValue("Short");

                // Auto‑fit columns to the content
                worksheet.AutoFitColumns();

                // Save the workbook
                string outputPath = "AutoFitColumnsDemo.xlsx";
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
            AutoFitColumnsExample.Run();
        }
    }
}