// Title: How to apply a custom number format for scientific notation with two decimal places using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell number format to scientific notation with two decimal places | C# custom number format 0.00E+00 using Aspose.Cells style flag | Apply only number format style to a single cell in an Aspose.Cells workbook | Create a style with scientific notation format in Aspose.Cells C# | Save an Excel file with scientific notation formatting using Aspose.Cells
// Tags: custom scientific notation format Aspose.Cells | apply number format style flag C# | 0.00E+00 custom format Aspose.Cells | single cell number format Aspose.Cells | save workbook with scientific notation Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, writes a numeric value to cell A1, defines a style with the custom format "0.00E+00" to display numbers in scientific notation with two decimal places, applies only the number‑format part of the style to the cell using a StyleFlag, and saves the file as ScientificNotation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a sample numeric value into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(12345.6789);

        // Create a style with a custom number format for scientific notation
        // "0.00E+00" displays the number in scientific notation with two decimal places
        Style sciStyle = workbook.CreateStyle();
        sciStyle.Custom = "0.00E+00";

        // Apply the style to the cell (only the number format part)
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;
        cell.SetStyle(sciStyle, flag);

        // Save the workbook to a file
        workbook.Save("ScientificNotation.xlsx");
    }
}
