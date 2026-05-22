using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add a single worksheet to the workbook
        // The Add() method returns the index of the newly added worksheet
        int newSheetIndex = workbook.Worksheets.Add();

        // Optionally set a name for the new worksheet
        workbook.Worksheets[newSheetIndex].Name = "MySheet";

        // Save the workbook to verify the result (optional)
        workbook.Save("MyWorkbook.xlsx");
    }
}