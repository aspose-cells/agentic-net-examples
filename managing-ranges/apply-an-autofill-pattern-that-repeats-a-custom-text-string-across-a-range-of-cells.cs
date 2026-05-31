using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class AutoFillCustomTextDemo
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
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the custom text to repeat
                string customText = "SampleText";

                // Fill the source range with the custom text (e.g., A1 and A2)
                cells["A1"].PutValue(customText);
                cells["A2"].PutValue(customText);

                // Create source and target ranges
                AsposeRange sourceRange = cells.CreateRange("A1:A2");
                AsposeRange targetRange = cells.CreateRange("A3:A10");

                // Use AutoFill to repeat the custom text across the target range
                sourceRange.AutoFill(targetRange, AutoFillType.Copy);

                // Save the workbook
                string outputPath = "AutoFillCustomTextDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}