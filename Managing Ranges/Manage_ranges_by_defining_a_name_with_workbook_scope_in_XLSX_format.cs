using Aspose.Cells;
using System;

class DefineWorkbookScopeName
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and give it a friendly name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate some sample data in the worksheet
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        sheet.Cells["B3"].PutValue(40);

        // Define a workbook‑scoped named range that refers to A1:B3
        int nameIndex = workbook.Worksheets.Names.Add("MyWorkbookRange");
        Name namedRange = workbook.Worksheets.Names[nameIndex];
        // The RefersTo formula must start with '=' and include the sheet name
        namedRange.RefersTo = $"={sheet.Name}!$A$1:$B$3";
        // SheetIndex = 0 indicates a global (workbook‑scoped) name
        namedRange.SheetIndex = 0;

        // Sort defined names before saving (optional but recommended for many names)
        workbook.Worksheets.SortNames();

        // Save the workbook in XLSX format
        workbook.Save("WorkbookScopeNamedRange.xlsx");
    }
}