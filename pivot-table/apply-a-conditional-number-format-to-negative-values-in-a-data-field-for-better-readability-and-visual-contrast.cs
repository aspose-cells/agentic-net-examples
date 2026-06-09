using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsConditionalNumberFormat
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (both positive and negative values)
                sheet.Cells["A1"].PutValue(1250.75);
                sheet.Cells["A2"].PutValue(-845.30);
                sheet.Cells["A3"].PutValue(4300);
                sheet.Cells["A4"].PutValue(-1200);
                sheet.Cells["A5"].PutValue(0);

                // Create a style with a custom number format.
                // The format shows positive numbers normally and negative numbers in red.
                Style customStyle = workbook.CreateStyle();
                customStyle.Custom = "_-€ #,##0.00;[Red]-€ #,##0.00";

                // Use a StyleFlag to apply only the number format (preserves other cell properties)
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Apply the style to the range containing the data
                AsposeRange dataRange = sheet.Cells.CreateRange("A1:A5");
                dataRange.ApplyStyle(customStyle, flag);

                // Define output file path
                string outputPath = "ConditionalNumberFormat.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}