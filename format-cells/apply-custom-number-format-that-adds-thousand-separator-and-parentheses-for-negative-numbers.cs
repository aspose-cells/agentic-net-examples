using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CustomNumberFormatDemo
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

                // Set a positive and a negative value to demonstrate the format
                sheet.Cells["A1"].PutValue(1234567.89);   // Positive number
                sheet.Cells["A2"].PutValue(-1234567.89); // Negative number

                // Create a style with a custom number format:
                // "#,##0;(#,##0)" adds thousand separators and encloses negative numbers in parentheses
                Style style = workbook.CreateStyle();
                style.Custom = "#,##0;(#,##0)";

                // Use StyleFlag to apply only the number format (leaving other style attributes unchanged)
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Apply the style to the range containing the two cells
                Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A2");
                range.ApplyStyle(style, flag);

                // Save the workbook to an XLSX file
                string filePath = "CustomNumberFormatDemo.xlsx";
                workbook.Save(filePath);

                // Optional: reload the file to verify the format was saved correctly
                if (File.Exists(filePath))
                {
                    Workbook verify = new Workbook(filePath);
                    Console.WriteLine("A1 format: " + verify.Worksheets[0].Cells["A1"].GetStyle().Custom);
                    Console.WriteLine("A2 format: " + verify.Worksheets[0].Cells["A2"].GetStyle().Custom);
                }
                else
                {
                    Console.WriteLine($"File not found: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}