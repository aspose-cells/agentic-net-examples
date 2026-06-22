using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the default column width for the entire worksheet to 15 characters
        worksheet.Cells.StandardWidth = 15;

        // Optional: display the set width
        Console.WriteLine("Standard column width set to: " + worksheet.Cells.StandardWidth);

        // Save the workbook
        workbook.Save("DefaultColumnWidth.xlsx");
    }
}