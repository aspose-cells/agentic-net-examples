using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExtractRawPercentageDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Put a numeric value that represents 50%
                Cell percentCell = cells["A1"];
                percentCell.PutValue(0.5);

                // Apply a built‑in percentage number format (ID 10)
                Style style = percentCell.GetStyle();
                style.Number = 10; // 10 = "0%" format
                percentCell.SetStyle(style);

                // Get the formatted string (e.g., "50%")
                string formatted = percentCell.StringValue;
                Console.WriteLine($"Formatted StringValue: {formatted}");

                // Get the raw value without any formatting using GetStringValue with None strategy
                string raw = percentCell.GetStringValue(CellValueFormatStrategy.None);
                Console.WriteLine($"Raw value via GetStringValue(None): {raw}");

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                string outputPath = "ExtractRawPercentageDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExtractRawPercentageDemo.Run();
        }
    }
}