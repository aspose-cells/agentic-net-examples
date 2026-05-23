using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a vertical header column (A1:A5) with sample text
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["A2"].PutValue("Header 2");
            sheet.Cells["A3"].PutValue("Header 3");
            sheet.Cells["A4"].PutValue("Header 4");
            sheet.Cells["A5"].PutValue("Header 5");

            // Create a style and set the text rotation angle to 45 degrees
            Style rotationStyle = workbook.CreateStyle();
            rotationStyle.RotationAngle = 45;

            // Create a style flag to indicate that the rotation setting should be applied
            StyleFlag flag = new StyleFlag();
            flag.Rotation = true;

            // Define the vertical header range (A1:A5) and apply the style with the flag
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(0, 0, 5, 1); // rows 0‑4, column 0 (A)
            headerRange.ApplyStyle(rotationStyle, flag);

            // Save the workbook
            string outputPath = "VerticalHeaderRotation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}