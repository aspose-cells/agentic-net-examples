using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace WorksheetScopedNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add two worksheets
            Worksheet dataSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            dataSheet.Name = "DataSheet";

            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "Summary";

            // Populate some data in DataSheet (A1:B3)
            dataSheet.Cells["A1"].PutValue("Item");
            dataSheet.Cells["B1"].PutValue("Qty");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("Orange");
            dataSheet.Cells["B3"].PutValue(15);

            // Create a worksheet‑scoped named range "DataRange" that belongs to DataSheet
            int nameIndex = workbook.Worksheets.Names.Add("DataRange");
            Name dataRangeName = workbook.Worksheets.Names[nameIndex];
            dataRangeName.RefersTo = $"={dataSheet.Name}!$A$1:$B$3";
            dataRangeName.SheetIndex = dataSheet.Index + 1; // one‑based index for worksheet‑scoped names

            // Retrieve the named range via the Name object
            Name retrievedName = workbook.Worksheets.Names["DataRange"];
            AsposeRange range = retrievedName.GetRange(); // returns the range on DataSheet

            // Demonstrate that the range knows its worksheet
            Console.WriteLine($"Range belongs to worksheet: {range.Worksheet.Name}");
            Console.WriteLine($"Range address: {range.RefersTo}");

            // Use the worksheet reference to modify a cell outside the original range
            range.Worksheet.Cells["C1"].PutValue("Total Qty");
            range.Worksheet.Cells["C2"].Formula = "=SUM(DataRange)";

            // Alternatively, obtain the range directly from the worksheet collection,
            // specifying the sheet index (zero‑based) and include tables flag = false
            AsposeRange directRange = workbook.Worksheets.GetRangeByName("DataRange", dataSheet.Index, false);
            Console.WriteLine($"Directly obtained range address: {directRange.RefersTo}");

            // Save the workbook
            workbook.Save("WorksheetScopedNamedRangeDemo.xlsx");
        }
    }
}