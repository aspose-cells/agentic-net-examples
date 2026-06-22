using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class ApplyHeaderFillColor
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the header range (first row, columns A to D)
            AsposeRange headerRange = worksheet.Cells.CreateRange("A1:D1");

            // Create a style for the header cells
            Style headerStyle = workbook.CreateStyle();

            // Set a custom RGB fill color (teal) and solid pattern
            headerStyle.ForegroundColor = Color.FromArgb(0, 128, 128);
            headerStyle.Pattern = BackgroundType.Solid;

            // Optional: set font color for better readability
            headerStyle.Font.Color = Color.White;

            // Create a StyleFlag to apply cell shading and font color
            StyleFlag flag = new StyleFlag
            {
                CellShading = true,
                FontColor = true
            };

            // Apply the style to the header range
            headerRange.ApplyStyle(headerStyle, flag);

            // Add sample header text
            worksheet.Cells["A1"].PutValue("Header 1");
            worksheet.Cells["B1"].PutValue("Header 2");
            worksheet.Cells["C1"].PutValue("Header 3");
            worksheet.Cells["D1"].PutValue("Header 4");

            // Save the workbook
            workbook.Save("HeaderFillColor.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}