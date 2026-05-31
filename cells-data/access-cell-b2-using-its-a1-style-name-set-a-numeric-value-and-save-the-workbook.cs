using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell B2 using A1 notation and set a numeric value
        worksheet.Cells["B2"].PutValue(123.45);

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}