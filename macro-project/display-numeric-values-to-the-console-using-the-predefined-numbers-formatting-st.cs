using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a style that uses a predefined number format.
        // Number = 2 corresponds to the "0.00" decimal format.
        Style numberStyle = workbook.CreateStyle();
        numberStyle.Number = 2;

        // Apply the style to the cells that will hold numeric values.
        worksheet.Cells["A1"].SetStyle(numberStyle);
        worksheet.Cells["A2"].SetStyle(numberStyle);
        worksheet.Cells["A3"].SetStyle(numberStyle);

        // Insert numeric values.
        worksheet.Cells["A1"].PutValue(123);
        worksheet.Cells["A2"].PutValue(45.678);
        worksheet.Cells["A3"].PutValue(0.5);

        // Output the formatted string values to the console.
        Console.WriteLine(worksheet.Cells["A1"].DisplayStringValue); // 123.00
        Console.WriteLine(worksheet.Cells["A2"].DisplayStringValue); // 45.68
        Console.WriteLine(worksheet.Cells["A3"].DisplayStringValue); // 0.50

        // Save the workbook (optional, demonstrates persistence).
        workbook.Save("NumberFormatDemo.xlsx");
    }
}