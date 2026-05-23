using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class ApplyStyleToNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in the range A1:C3
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["C1"].PutValue("Header3");
            cells["A2"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["C2"].PutValue(30);
            cells["A3"].PutValue(40);
            cells["B3"].PutValue(50);
            cells["C3"].PutValue(60);

            // Create a range covering A1:C3 and assign a name to it
            Aspose.Cells.Range namedRange = cells.CreateRange("A1", "C3");
            namedRange.Name = "MyRange";

            // Define a custom style to be applied to the entire range
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Name = "Calibri";
            customStyle.Font.Size = 12;
            customStyle.Font.IsBold = true;
            customStyle.Font.Color = Color.White;
            customStyle.ForegroundColor = Color.DarkBlue;
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.HorizontalAlignment = TextAlignmentType.Center;
            customStyle.VerticalAlignment = TextAlignmentType.Center;

            // Apply the custom style to the named range
            namedRange.SetStyle(customStyle);

            // Define output file path
            string outputPath = "StyledNamedRange.xlsx";

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Log or display the error
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}