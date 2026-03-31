using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

public class CopyRangeExample
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet and set it as the source sheet
        Worksheet sheetSrc = workbook.Worksheets[0];
        sheetSrc.Name = "Source";

        // Populate the source range with sample data
        sheetSrc.Cells["A1"].PutValue("Source Data");
        sheetSrc.Cells["B1"].PutValue(100);
        sheetSrc.Cells["A2"].PutValue("More Data");
        sheetSrc.Cells["B2"].PutValue(200);

        // Add a second worksheet to serve as the destination sheet
        Worksheet sheetDest = workbook.Worksheets[workbook.Worksheets.Add()];
        sheetDest.Name = "Destination";

        // Create source and destination Range objects
        AsposeRange rangeSrc = sheetSrc.Cells.CreateRange("A1:B2");
        AsposeRange rangeDest = sheetDest.Cells.CreateRange("C3:D4");

        // Copy data (including formulas) from the source range to the destination range
        rangeDest.CopyData(rangeSrc);

        // Save the workbook to an XLSX file
        workbook.Save("RangeCopyDataDemo.xlsx");
    }
}