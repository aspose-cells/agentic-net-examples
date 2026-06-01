using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class ApplyCustomNumberFormatToColumnC
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a range starting at C3 (row index 2, column index 2)
            AsposeRange offsetRange = cells.CreateRange(2, 2, 1, 1);

            // Get the entire column that contains the offset range (column C)
            AsposeRange entireColumnC = offsetRange.EntireColumn;

            // Define a custom number format (percentage with two decimals)
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "0.00%";

            // Apply only the number format
            StyleFlag flag = new StyleFlag { NumberFormat = true };

            // Apply the custom number format to the whole column C
            entireColumnC.ApplyStyle(customStyle, flag);

            // Save the workbook
            string outputPath = "ColumnC_CustomNumberFormat.xlsx";
            workbook.Save(outputPath);
        }
    }
}