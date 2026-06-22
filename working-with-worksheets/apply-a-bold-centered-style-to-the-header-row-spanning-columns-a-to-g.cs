using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style: bold font and centered horizontally
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;

            // Define which style properties should be applied
            StyleFlag styleFlag = new StyleFlag
            {
                FontBold = true,
                HorizontalAlignment = true
            };

            // Create a range that covers cells A1 to G1 (first row, columns A‑G)
            Aspose.Cells.Range headerRange = worksheet.Cells.CreateRange("A1:G1");

            // Apply the style to the defined range
            headerRange.ApplyStyle(headerStyle, styleFlag);

            // Optional: autofit the header row height
            worksheet.AutoFitRow(0);

            // Save the workbook
            string outputPath = "HeaderStyled.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}