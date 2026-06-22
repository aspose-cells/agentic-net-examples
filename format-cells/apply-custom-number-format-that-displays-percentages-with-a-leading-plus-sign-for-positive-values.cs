using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class PercentageWithPlusSignDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put a sample numeric value (e.g., 12.34%)
                Cell cell = sheet.Cells["A1"];
                cell.PutValue(0.1234); // 12.34%

                // Create a style and set a custom number format:
                // Positive values: +0.00%
                // Negative values: -0.00%
                // Zero values: 0%
                Style style = workbook.CreateStyle();
                style.Custom = "+0.00%;-0.00%;0%";

                // Apply the style to the cell
                cell.SetStyle(style);

                // Define output file path
                string outputPath = "PercentageWithPlusSignDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error in PercentageWithPlusSignDemo: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                PercentageWithPlusSignDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}