using System;
using Aspose.Cells;

class RepeatRowsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure rows 1 to 3 to repeat on every printed page
        // The range uses absolute A1 notation
        worksheet.PageSetup.PrintTitleRows = "$1:$3";

        // Save the workbook
        workbook.Save("RepeatRowsDemo.xlsx");
    }
}