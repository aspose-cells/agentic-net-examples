using System;
using Aspose.Cells;

class RemoveWorksheetsByIndex
{
    static void Main()
    {
        // Create a new workbook (default contains one worksheet)
        Workbook wb = new Workbook();

        // Add additional worksheets
        wb.Worksheets.Add("Sheet1");
        wb.Worksheets.Add("Sheet2");
        wb.Worksheets.Add("Sheet3");

        // Display the original count
        Console.WriteLine("Original worksheet count: " + wb.Worksheets.Count);

        // Remove the worksheet at index 1 (the second worksheet)
        wb.Worksheets.RemoveAt(1);

        // Display the count after removal
        Console.WriteLine("Worksheet count after removal: " + wb.Worksheets.Count);

        // Save the workbook to verify the result
        wb.Save("RemovedSheetDemo.xlsx");
    }
}