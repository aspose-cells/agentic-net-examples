using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace NamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add source and destination worksheets
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // Populate source sheet with sample data (A1:C3)
            sourceSheet.Cells["A1"].PutValue("Item");
            sourceSheet.Cells["B1"].PutValue("Qty");
            sourceSheet.Cells["C1"].PutValue("Price");
            sourceSheet.Cells["A2"].PutValue("Apple");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["C2"].PutValue(0.5);
            sourceSheet.Cells["A3"].PutValue("Banana");
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["C3"].PutValue(0.3);

            // Create a named range "SourceData" that refers to A1:C3 on the source sheet
            int nameIndex = workbook.Worksheets.Names.Add("SourceData");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            namedRange.RefersTo = $"={sourceSheet.Name}!$A$1:$C$3";

            // Retrieve the range using Name.GetRange()
            AsposeRange srcRange = namedRange.GetRange();

            // Demonstrate accessing the same range via WorksheetCollection.GetRangeByName
            AsposeRange srcRangeViaCollection = workbook.Worksheets.GetRangeByName("SourceData");
            Console.WriteLine($"Range address via GetRange: {srcRange.Address}");
            Console.WriteLine($"Range address via GetRangeByName: {srcRangeViaCollection.Address}");

            // Create a destination range with the same size on the destination sheet
            AsposeRange destRange = destSheet.Cells.CreateRange("A1", "C3");

            // Copy values from source range to destination range
            destRange.CopyValue(srcRange);

            // Verify copy by printing a value from the destination sheet
            Console.WriteLine($"Copied value at Destination!A2: {destSheet.Cells["A2"].StringValue}");

            // Save the workbook
            workbook.Save("NamedRangeCopyDemo.xlsx");
        }
    }
}