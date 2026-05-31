using System;
using System.IO;
using Aspose.Cells;

class ApplyCustomNumberFormat
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (positive, negative and zero)
            sheet.Cells["A1"].PutValue(1234);
            sheet.Cells["A2"].PutValue(-5678);
            sheet.Cells["A3"].PutValue(0);

            // Create a style with the desired custom number format
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0;[Red]-#,##0";

            // Apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Apply the style to the range A1:A3 (rows 0‑2, column 0)
            Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, 3, 1);
            range.ApplyStyle(customStyle, flag);

            // Save the workbook
            string outputPath = "CustomNumberFormat.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}