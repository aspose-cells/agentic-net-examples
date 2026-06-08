using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the default style's custom number format to "dd-mmm-yyyy"
        Style defaultStyle = wb.DefaultStyle;
        defaultStyle.Custom = "dd-mmm-yyyy";
        wb.DefaultStyle = defaultStyle;

        // Demonstrate the format on a sample date cell
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue(DateTime.Now);
        ws.Cells["A1"].SetStyle(defaultStyle);

        // Save the workbook
        wb.Save("DefaultDateFormat.xlsx");
    }
}