using System;
using Aspose.Cells;

class ListWorkbookScopedNames
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a workbook‑scoped named range
        int globalIndex = workbook.Worksheets.Names.Add("GlobalRange");
        workbook.Worksheets.Names[globalIndex].RefersTo = "=Sheet1!$A$1:$A$5";

        // Add a worksheet‑scoped named range (for demonstration)
        int sheetIndex = workbook.Worksheets.Names.Add("SheetRange");
        workbook.Worksheets.Names[sheetIndex].RefersTo = "=Sheet1!$B$1:$B$5";
        workbook.Worksheets.Names[sheetIndex].SheetIndex = 0; // scope to first worksheet

        // Retrieve only the workbook‑scoped names
        Name[] workbookScopedNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);

        // Output the names to the console
        Console.WriteLine("Workbook‑scoped named ranges:");
        foreach (Name name in workbookScopedNames)
        {
            Console.WriteLine(name.Text);
        }

        // Save the workbook (optional)
        workbook.Save("WorkbookScopedNames.xlsx");
    }
}