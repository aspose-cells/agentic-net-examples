using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class GetRawNumericStringDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Put a numeric value into cell A1
                cells["A1"].PutValue(12345.6789);

                // Retrieve the raw numeric string without any formatting
                // CellValueFormatStrategy.None means no formatting is applied
                string rawNumericString = cells["A1"].GetStringValue(CellValueFormatStrategy.None);

                // Display the result
                Console.WriteLine("Raw numeric string (no formatting): " + rawNumericString);

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                string outputPath = "GetRawNumericStringDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            GetRawNumericStringDemo.Run();
        }
    }
}