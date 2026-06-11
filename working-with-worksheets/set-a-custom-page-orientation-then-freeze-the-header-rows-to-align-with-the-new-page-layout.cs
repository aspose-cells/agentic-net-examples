using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set custom page orientation (Landscape)
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Add sample data with a header row
        worksheet.Cells["A1"].PutValue("Header");
        for (int i = 2; i <= 30; i++)
        {
            worksheet.Cells[$"A{i}"].PutValue($"Data row {i - 1}");
        }

        // Freeze the header row so it stays visible when scrolling
        // Freeze at row index 1 (second row) with 1 frozen row and 0 frozen columns
        worksheet.FreezePanes(1, 0, 1, 0);

        // Ensure the header row repeats on each printed page
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // Save the workbook
        workbook.Save("CustomOrientation_FrozenHeader.xlsx");
    }
}