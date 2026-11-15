// Title: Repeat header row and first column on every printed page using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# repeat header row on each printed page | How to set print title columns in Aspose.Cells .NET workbook | Print Excel worksheet with repeated first column as title using Aspose.Cells | Fit worksheet to one page wide while repeating rows and columns with Aspose.Cells | C# Aspose.Cells page setup repeat rows and columns on print
// Tags: Aspose.Cells PrintTitleRows property C# | Aspose.Cells PrintTitleColumns setting | Aspose.Cells page setup repeat rows columns | Aspose.Cells fit worksheet to one page wide | Aspose.Cells repeat header and first column on print

using Aspose.Cells;
using System;

// The example creates a workbook, adds header values in the first row and titles in the first column, then uses the PageSetup properties PrintTitleRows = "$1:$1" and PrintTitleColumns = "$A:$A" to repeat the header row and first column on every printed page, fits the sheet to one page wide, and saves the file as RepeatedHeaders.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Populate sample data -----
        // Header row (will be repeated on each printed page)
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");

        // First column titles (will be repeated on each printed page) and sample data
        for (int row = 2; row <= 20; row++)
        {
            sheet.Cells[row - 1, 0].PutValue("Item " + (row - 1)); // First column title
            sheet.Cells[row - 1, 1].PutValue("Name " + (row - 1));
            sheet.Cells[row - 1, 2].PutValue(row * 10);
        }

        // ----- Set print options -----
        // Repeat the header row (first row) on every printed page
        sheet.PageSetup.PrintTitleRows = "$1:$1";

        // Repeat the first column on every printed page
        sheet.PageSetup.PrintTitleColumns = "$A:$A";

        // Optional: fit the sheet to one page wide, height can span multiple pages
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0;

        // Save the workbook
        workbook.Save("RepeatedHeaders.xlsx", SaveFormat.Xlsx);
    }
}
