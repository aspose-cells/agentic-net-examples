using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – copy a worksheet and place it directly after the original
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add some data
        Worksheet originalSheet = workbook.Worksheets[0];
        originalSheet.Cells["A1"].PutValue("Original Worksheet");

        // Add a copy of the first worksheet (adds at the end of the collection)
        int copiedSheetIndex = workbook.Worksheets.AddCopy(0);

        // Move the copied worksheet so it is placed directly after the original sheet (index 1)
        workbook.Worksheets[copiedSheetIndex].MoveTo(1);

        // Optionally rename the copied sheet
        workbook.Worksheets[1].Name = "Copied Worksheet";

        // Save the workbook
        workbook.Save("CopyAfterOriginal.xlsx");
    }
}