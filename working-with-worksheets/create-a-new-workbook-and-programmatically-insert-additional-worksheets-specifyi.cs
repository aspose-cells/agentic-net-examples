using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (default workbook contains one worksheet named "Sheet1")
        Workbook workbook = new Workbook();

        // Insert a worksheet at position 0 and give it a custom name
        Worksheet introSheet = workbook.Worksheets.Insert(0, SheetType.Worksheet, "Intro");
        introSheet.Cells["A1"].PutValue("Welcome to the Intro sheet");

        // Insert another worksheet at position 2 (after the existing sheets)
        // At this point the workbook has: Intro (0), Sheet1 (1)
        Worksheet dataSheet = workbook.Worksheets.Insert(2, SheetType.Worksheet, "Data");
        dataSheet.Cells["A1"].PutValue(123);
        dataSheet.Cells["B1"].PutValue(456);

        // Add a worksheet at the end with a specific name
        Worksheet summarySheet = workbook.Worksheets.Add("Summary");
        summarySheet.Cells["A1"].PutValue("Summary information");

        // Save the workbook to a file
        workbook.Save("InsertedSheets.xlsx");
    }
}