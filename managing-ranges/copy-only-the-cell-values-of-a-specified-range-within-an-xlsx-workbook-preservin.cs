using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define source and destination ranges
        Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("A1:C3");
        Aspose.Cells.Range destinationRange = sheet.Cells.CreateRange("D1:F3");

        // Copy only the cell values
        destinationRange.CopyValue(sourceRange);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}