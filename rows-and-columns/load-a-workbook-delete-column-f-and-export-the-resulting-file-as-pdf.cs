using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx"); // replace with your source file path

        // Delete column F (zero‑based index 5) from the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.DeleteColumn(5);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf); // replace with desired output path
    }
}