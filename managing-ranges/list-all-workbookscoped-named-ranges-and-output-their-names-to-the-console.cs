using Aspose.Cells;
using System;

class ListWorkbookScopedNames
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a workbook‑scoped named range
        int globalIdx = workbook.Worksheets.Names.Add("GlobalRange");
        workbook.Worksheets.Names[globalIdx].RefersTo = "=Sheet1!$A$1:$A$5";

        // Add a worksheet‑scoped named range (for demonstration)
        int sheetIdx = workbook.Worksheets.Names.Add("SheetRange");
        workbook.Worksheets.Names[sheetIdx].RefersTo = "=Sheet1!$B$1:$B$5";
        workbook.Worksheets.Names[sheetIdx].SheetIndex = 0; // 0 = first worksheet

        // Retrieve only workbook‑scoped names
        Name[] workbookScopedNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);

        Console.WriteLine("Workbook‑scoped named ranges:");
        if (workbookScopedNames != null && workbookScopedNames.Length > 0)
        {
            foreach (Name name in workbookScopedNames)
            {
                Console.WriteLine(name.Text);
            }
        }
        else
        {
            Console.WriteLine("No workbook‑scoped named ranges found.");
        }

        // Save the workbook (optional)
        workbook.Save("WorkbookScopedNames.xlsx");
    }
}