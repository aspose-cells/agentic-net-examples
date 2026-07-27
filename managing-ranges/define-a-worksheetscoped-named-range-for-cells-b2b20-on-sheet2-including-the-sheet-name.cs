using System;
using Aspose.Cells;

class WorksheetScopedNamedRange
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a worksheet named "Sheet2"
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Add a new name to the workbook's name collection
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        Name namedRange = workbook.Worksheets.Names[nameIndex];

        // Define the range B2:B20 on Sheet2
        namedRange.RefersTo = "=Sheet2!$B$2:$B$20";

        // Set the scope to Sheet2 (one‑based sheet index = 2)
        namedRange.SheetIndex = 2;

        // Save the workbook
        workbook.Save("WorksheetScopedNamedRange.xlsx");
    }
}