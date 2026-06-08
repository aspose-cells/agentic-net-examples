using System;
using System.IO;
using Aspose.Cells;
using Range = Aspose.Cells.Range;   // Resolve ambiguity with System.Range

class ApplyCustomPercentageFormat
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with positive and negative decimal values (percentages)
            sheet.Cells["A1"].PutValue(0.25);    // 25%
            sheet.Cells["A2"].PutValue(-0.125); // -12.5%

            // Create a style with a custom number format:
            // Positive percentages display normally, negative percentages appear in red with a minus sign
            Style style = workbook.CreateStyle();
            style.Custom = "0.00%;[Red]-0.00%";

            // Apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the target range (A1:A2)
            Range range = sheet.Cells.CreateRange("A1", "A2");
            range.ApplyStyle(style, flag);

            // Ensure output directory exists
            string outputPath = "CustomPercentageFormat.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}