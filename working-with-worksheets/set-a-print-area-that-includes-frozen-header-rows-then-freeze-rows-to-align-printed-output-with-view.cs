using System;
using Aspose.Cells;

class PrintAreaAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add header row
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");

        // Add sample data rows
        for (int i = 2; i <= 30; i++)
        {
            worksheet.Cells[$"A{i}"].PutValue($"Row{i - 1} Col1");
            worksheet.Cells[$"B{i}"].PutValue($"Row{i - 1} Col2");
        }

        // Set the print area to include the header and all data rows
        worksheet.PageSetup.PrintArea = "A1:B30";

        // Repeat the header row on each printed page
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // Freeze the header row so it stays visible while scrolling
        // Freeze at cell A2, freezing 1 row (the header) and 0 columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("PrintAreaAndFreezeDemo.xlsx");
    }
}