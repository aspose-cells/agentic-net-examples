using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a new style using the Workbook.CreateStyle method (rule)
        Style fontStyle = workbook.CreateStyle();

        // Set the desired font family and size
        fontStyle.Font.Name = "Calibri";
        fontStyle.Font.Size = 14;

        // Apply the style to specific cells
        cells["A1"].PutValue("Styled A1");
        cells["A1"].SetStyle(fontStyle);

        cells["B2"].PutValue("Styled B2");
        cells["B2"].SetStyle(fontStyle);

        cells["C3"].PutValue("Styled C3");
        cells["C3"].SetStyle(fontStyle);

        // Save the workbook (lifecycle rule)
        workbook.Save("StyledCells.xlsx");
    }
}