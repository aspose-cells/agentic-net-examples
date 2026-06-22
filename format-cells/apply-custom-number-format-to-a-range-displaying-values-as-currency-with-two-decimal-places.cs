using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ApplyCurrencyNumberFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample numeric values to the cells
                sheet.Cells["A1"].PutValue(1234.5);
                sheet.Cells["A2"].PutValue(5678.9);
                sheet.Cells["A3"].PutValue(0.75);

                // Create a style with a custom currency format (two decimal places)
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Custom = "$#,##0.00";

                // Configure a StyleFlag to apply only the number format part of the style
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Define the range to which the style will be applied (A1:A3)
                Aspose.Cells.Range targetRange = sheet.Cells.CreateRange("A1", "A3");
                targetRange.ApplyStyle(currencyStyle, flag);

                // Save the workbook to a file
                string outputPath = "CurrencyNumberFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCurrencyNumberFormat.Run();
        }
    }
}