using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (macro‑enabled by default)
        Workbook workbook = new Workbook();

        // Remove any VBA/macros to make it macro‑free
        workbook.RemoveMacro();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Hide zero values in the worksheet
        sheet.DisplayZeros = false;

        // Set a custom top margin (value is in inches)
        sheet.PageSetup.TopMargin = 0.75; // example: 0.75 inches

        // Save the workbook as a CSV file
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}