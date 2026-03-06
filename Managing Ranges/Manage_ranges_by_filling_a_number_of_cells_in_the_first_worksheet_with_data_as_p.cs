using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class ManageRangesExample
{
    static void Main()
    {
        // Create a new workbook and get the cells collection of the first worksheet
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Populate the source range (A1:A2) with initial numeric data
        cells["A1"].PutValue(1);
        cells["A2"].PutValue(2);

        // Define the source range that will be used for autofill
        AsposeRange sourceRange = cells.CreateRange("A1", "A2");

        // Define the target range where the series will be extended (A3:A20)
        AsposeRange targetRange = cells.CreateRange("A3", "A20");

        // Autofill the target range using a series (1,2,3,...)
        sourceRange.AutoFill(targetRange, AutoFillType.Series);

        // Save the workbook in XLSX format
        workbook.Save("ManagedRanges.xlsx");
    }
}