using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace AsposeCellsExamples
{
    public class CustomNumberFormatDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put sample values: a positive and a negative number
            sheet.Cells["A1"].PutValue(1234.56);
            sheet.Cells["A2"].PutValue(-1234.56);

            // Create a style and set a custom number format.
            // Positive numbers: normal; negative numbers: red and shown in parentheses.
            Style style = workbook.CreateStyle();
            style.Custom = "#,##0;[Red](#,##0)";

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the range A1:A2
            AsposeRange range = sheet.Cells.CreateRange("A1", "A2");
            range.ApplyStyle(style, flag);

            // Save the workbook
            string outputPath = "CustomNumberFormatDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}