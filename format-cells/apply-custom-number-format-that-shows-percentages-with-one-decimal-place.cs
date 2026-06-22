using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class CustomPercentageFormatDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with fractional values (e.g., 0.1234 = 12.34%)
                sheet.Cells["A1"].PutValue(0.1234);
                sheet.Cells["A2"].PutValue(0.5678);
                sheet.Cells["A3"].PutValue(0.9);

                // Define a custom percentage style with one decimal place
                Style percentStyle = workbook.CreateStyle();
                percentStyle.Custom = "0.0%"; // e.g., 12.3%

                // Apply the style to the range A1:A3
                AsposeRange range = sheet.Cells.CreateRange("A1:A3");
                range.ApplyStyle(percentStyle, new StyleFlag { NumberFormat = true });

                // Save the workbook
                string outputPath = "CustomPercentageFormatDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}