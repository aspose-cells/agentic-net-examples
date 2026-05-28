using System;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all pivot tables from every worksheet in the workbook
        workbook.Worksheets.ClearPivottables();

        // Save the modified workbook to a new file (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}