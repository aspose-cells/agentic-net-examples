using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – split semicolon‑delimited text into columns
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate column A with semicolon‑separated values
        sheet.Cells["A1"].PutValue("John;Doe;30");
        sheet.Cells["A2"].PutValue("Jane;Smith;28");

        // Set up text load options to use semicolon as the delimiter
        TxtLoadOptions options = new TxtLoadOptions();
        options.Separator = ';'; // semicolon delimiter

        // Split the text in column A (row 0, column 0) for the first 2 rows
        sheet.Cells.TextToColumns(0, 0, 2, options);

        // Verify the split results (optional)
        Console.WriteLine("B1: " + sheet.Cells["B1"].StringValue); // Expected: Doe
        Console.WriteLine("C2: " + sheet.Cells["C2"].StringValue); // Expected: 28

        // Save the workbook
        workbook.Save("SemicolonSplit.xlsx");
    }
}