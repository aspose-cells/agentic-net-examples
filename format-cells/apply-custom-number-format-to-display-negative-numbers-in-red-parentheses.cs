using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class NegativeNumberRedParenthesesDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a positive and a negative value in cells A1 and A2
            sheet.Cells["A1"].PutValue(1234.56);
            sheet.Cells["A2"].PutValue(-1234.56);

            // Create a style with custom number format:
            // Positive numbers: normal display
            // Negative numbers: red color and enclosed in parentheses
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0.00;[Red](#,##0.00)";

            // Apply only the number format part of the style
            StyleFlag styleFlag = new StyleFlag();
            styleFlag.NumberFormat = true;

            // Apply the style to the range containing the two cells
            AsposeRange range = sheet.Cells.CreateRange("A1:A2");
            range.ApplyStyle(style, styleFlag);

            // Define output file path
            string outputPath = "NegativeNumberRedParenthesesDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}