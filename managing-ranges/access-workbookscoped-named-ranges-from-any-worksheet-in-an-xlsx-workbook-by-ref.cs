using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class AccessWorkbookScopedNamedRanges
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Add worksheets
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Populate some data in Sheet1
        sheet1.Cells["A1"].PutValue("Item");
        sheet1.Cells["B1"].PutValue(10);
        sheet1.Cells["A2"].PutValue("Item2");
        sheet1.Cells["B2"].PutValue(20);

        // Create a workbook‑scoped (global) named range
        // The name is defined on a worksheet but is global by default
        sheet1.Cells.CreateRange("A1:B2").Name = "GlobalRange";

        // Access the named range from any worksheet using WorksheetCollection.GetRangeByName
        AsposeRange namedRange = workbook.Worksheets.GetRangeByName("GlobalRange");

        if (namedRange != null)
        {
            Console.WriteLine($"Named range found: {namedRange.Name}");
            Console.WriteLine($"Address: {namedRange.RefersTo}");
            // Example: read the first cell value of the range
            Console.WriteLine($"First cell value: {namedRange[0, 0].StringValue}");
        }
        else
        {
            Console.WriteLine("Named range not found.");
        }

        // Save the workbook (save rule)
        workbook.Save("WorkbookScopedNamedRanges.xlsx");
    }
}