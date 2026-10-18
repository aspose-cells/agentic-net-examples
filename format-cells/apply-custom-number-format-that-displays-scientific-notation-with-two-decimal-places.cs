// Title: How to apply a custom number format for scientific notation with two decimal places using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# format cell as scientific notation with two decimal places | set custom number format 0.00E+00 in Excel workbook using Aspose.Cells | C# example to display numeric value in scientific notation in Aspose.Cells | apply scientific notation style to a cell in Aspose.Cells for .NET | save workbook with scientific notation formatting using Aspose.Cells C#
// Tags: custom number format scientific notation Aspose.Cells | C# apply 0.00E+00 format Excel cell | Aspose.Cells style cell scientific notation | save workbook with scientific notation format .NET | format numeric value as scientific notation Aspose

using System;
using Aspose.Cells;

// The example creates a new workbook, writes the value 123456789 to cell A1, applies the custom number format "0.00E+00" to display the value in scientific notation with two decimal places, and saves the file as ScientificNotation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get a reference to cell A1
        Cell cell = sheet.Cells["A1"];

        // Set a numeric value that will be displayed in scientific notation
        cell.PutValue(123456789);

        // Create a style object
        Style style = cell.GetStyle();

        // Apply a custom number format for scientific notation with two decimal places
        // Format: 0.00E+00 (e.g., 1.23E+08)
        style.Custom = "0.00E+00";

        // Assign the style back to the cell
        cell.SetStyle(style);

        // Save the workbook (lifecycle: save)
        workbook.Save("ScientificNotation.xlsx");
    }
}
