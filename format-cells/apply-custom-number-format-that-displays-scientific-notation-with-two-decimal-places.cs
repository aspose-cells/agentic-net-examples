using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class ApplyScientificNumberFormat
    {
        public static void Main(string[] args)
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

            // Put a numeric value into cell A1
            sheet.Cells["A1"].PutValue(12345.6789);

            // Create a style with a custom scientific notation format (two decimal places)
            Style sciStyle = workbook.CreateStyle();
            sciStyle.Custom = "0.00E+00";

            // Configure a StyleFlag to apply only the number format
            StyleFlag flag = new StyleFlag
            {
                NumberFormat = true
            };

            // Apply the style to cell A1 using the flag
            AsposeRange range = sheet.Cells.CreateRange("A1");
            range.ApplyStyle(sciStyle, flag);

            // Define output file path
            string outputPath = "ScientificNumberFormat.xlsx";

            // Save the workbook to a file (overwrite if exists)
            workbook.Save(outputPath);
        }
    }
}