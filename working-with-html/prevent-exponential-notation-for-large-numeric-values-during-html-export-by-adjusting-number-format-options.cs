using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Large numeric value that would normally be rendered in exponential notation
        double largeValue = 1234567890123456; // 1.23456789012346E+15

        // Put the value into a cell
        sheet.Cells["A1"].PutValue(largeValue);

        // Apply a custom number format that forces plain decimal representation
        Style style = sheet.Cells["A1"].GetStyle();
        style.Custom = "0";               // No decimal places, no exponent
        sheet.Cells["A1"].SetStyle(style);

        // Ensure the global significant digits handling keeps full precision
        CellsHelper.SignificantDigitsType = SignificantDigitsType.G17;

        // Configure HTML save options (default options are sufficient for this case)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook as HTML; the large number will appear without exponential notation
        workbook.Save("LargeNumber.html", htmlOptions);
    }
}