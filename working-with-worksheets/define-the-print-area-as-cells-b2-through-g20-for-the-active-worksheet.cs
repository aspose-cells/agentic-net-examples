using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the active worksheet (first worksheet by default)
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the print area from B2 to G20
        worksheet.PageSetup.PrintArea = "B2:G20";

        // Save the workbook (optional, demonstrates that the setting is applied)
        workbook.Save("PrintAreaDemo.xlsx");
    }
}