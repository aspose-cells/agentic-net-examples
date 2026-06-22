using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class MixedNumberFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put a numeric value that will be displayed as a mixed fraction (e.g., 2.75 = 2 3/4)
                Cell cell = sheet.Cells["A1"];
                cell.PutValue(2.75);

                // Create a style with a custom number format for mixed fractions
                // "# ?/?" displays the integer part and the fractional part as a mixed number
                Style mixedFractionStyle = workbook.CreateStyle();
                mixedFractionStyle.Custom = "# ?/?";

                // Apply only the number format part of the style
                StyleFlag flag = new StyleFlag { NumberFormat = true };

                // Apply the style to the target cell (range of one cell)
                AsposeRange range = sheet.Cells.CreateRange("A1");
                range.ApplyStyle(mixedFractionStyle, flag);

                // Save the workbook
                string outputPath = "MixedNumberFormatDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main()
        {
            Run();
        }
    }
}