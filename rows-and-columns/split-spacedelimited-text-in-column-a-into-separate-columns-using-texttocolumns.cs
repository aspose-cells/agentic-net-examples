using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample space‑delimited data in column A
        sheet.Cells["A1"].PutValue("John Doe 30");
        sheet.Cells["A2"].PutValue("Jane Smith 28");
        sheet.Cells["A3"].PutValue("Bob Johnson 45");

        // Set up TextToColumns options to split on space
        TxtLoadOptions options = new TxtLoadOptions();
        options.Separator = ' '; // space character as delimiter

        // Split the text in column A (row 0, column 0) for 3 rows
        sheet.Cells.TextToColumns(0, 0, 3, options);

        // Optional verification
        Console.WriteLine(sheet.Cells["B1"].StringValue); // Expected: "Doe"
        Console.WriteLine(sheet.Cells["C2"].StringValue); // Expected: "28"

        // Save the workbook (lifecycle rule)
        workbook.Save("output.xlsx");
    }
}

// Author: Aspose.Cells .NET example – splits space‑delimited text using TextToColumns.