using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class NamedRangeWorkbookScopeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and give it a friendly name
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Populate some data in Sheet1
        sheet1.Cells["A1"].PutValue("Apple");
        sheet1.Cells["B1"].PutValue(10);
        sheet1.Cells["A2"].PutValue("Banana");
        sheet1.Cells["B2"].PutValue(20);

        // Create a global named range that covers A1:B2 in Sheet1
        // The range is workbook‑scoped because we do not set a sheet index on the Name object
        sheet1.Cells.CreateRange("A1:B2").Name = "FruitData";

        // Add a second worksheet to demonstrate accessing the named range from another sheet
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].PutValue("Values copied from the named range will appear below:");

        // Retrieve the named range using the workbook‑wide GetRangeByName method
        AsposeRange globalRange = workbook.Worksheets.GetRangeByName("FruitData");

        if (globalRange != null)
        {
            // Define a destination range in Sheet2 (C1:D2) and copy the values from the named range
            AsposeRange destination = sheet2.Cells.CreateRange("C1", "D2");
            destination.CopyValue(globalRange);
        }

        // Save the workbook in XLSX format
        workbook.Save("NamedRangeWorkbookScope.xlsx");
    }
}