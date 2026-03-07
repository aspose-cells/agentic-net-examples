using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the custom number format (e.g., US locale with two decimal places)
        string customNumberFormat = "[$-409]#,##0.00";

        // Create a style that uses the custom number format
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = customNumberFormat;

        // Prepare a StyleFlag to apply only the number format part of the style
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the style to the desired range.
        // Example: apply to the entire column B (index 1) that contains data.
        int firstRow = 0;
        int lastRow = worksheet.Cells.MaxDataRow; // last row with data
        if (lastRow < 0) lastRow = 0; // handle empty sheet

        Aspose.Cells.Range range = worksheet.Cells.CreateRange(firstRow, 1, lastRow - firstRow + 1, 1);
        range.ApplyStyle(customStyle, flag);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}