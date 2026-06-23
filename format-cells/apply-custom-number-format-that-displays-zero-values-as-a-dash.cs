using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class ZeroAsDashDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];

                // Ensure zero values are displayed (default is true)
                sheet.DisplayZeros = true;

                // Populate some cells with positive, negative and zero values
                sheet.Cells["A1"].PutValue(123);
                sheet.Cells["A2"].PutValue(0);
                sheet.Cells["A3"].PutValue(-45);
                sheet.Cells["A4"].PutValue(0);

                // Create a style with a custom number format:
                //   Positive numbers: 0
                //   Negative numbers: -0
                //   Zero values:      "-" (dash)
                //   Text values:      unchanged (@)
                Style style = wb.CreateStyle();
                style.Custom = "0;-0;\"-\";@";

                // Apply only the number format part of the style to the target range
                StyleFlag flag = new StyleFlag { NumberFormat = true };
                AsposeRange range = sheet.Cells.CreateRange("A1:A4");
                range.ApplyStyle(style, flag);

                // Save the workbook
                string outputPath = "ZeroAsDashDemo.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during ZeroAsDashDemo execution: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ZeroAsDashDemo.Run();
        }
    }
}